import {
	App,
	Component,
	MarkdownRenderer,
	Notice,
	base64ToArrayBuffer,
	requestUrl,
	setIcon,
} from 'obsidian';
import type { AiFormulaProvider } from './formulas';
import type { Scalar } from './model';
import type { PluginSettings } from './settings';

export type AiProvider = AiFormulaProvider;

export const AI_PROVIDER_LABELS: Record<AiProvider, string> = {
	claude: 'Claude',
	gemini: 'Gemini',
};

/** Shown in a cell while its AI formula waits for the answer. */
export const AI_LOADING_TEXT = 'Loading…';
/** Shown in a cell whose AI formula failed (the reason goes to a notice). */
export const AI_ERROR_TEXT = '#AI!';

/** How many AI formula requests run at the same time; the rest wait their turn. */
const MAX_PARALLEL_FORMULAS = 4;
/** Most answers kept in a file; the oldest go first. */
const MAX_CACHED_RESULTS = 2000;
/** Most model turns in one assistant reply (each tool round is one). */
const MAX_AGENT_STEPS = 16;

/* ------------------------------------------------------------------ */
/* Keys and requests                                                   */
/* ------------------------------------------------------------------ */

/** Where the API key of a provider is kept in Obsidian's secret storage. */
export function aiKeySecretId(provider: AiProvider): string {
	return `free-excel-${provider}-api-key`;
}

/** The API key of a provider, from Obsidian's secret storage; null when none is set. */
export function aiApiKey(
	app: App,
	settings: PluginSettings,
	provider: AiProvider,
): string | null {
	// Older versions let the user pick a secret by name: still read it when there is no key.
	const legacy =
		provider === 'claude'
			? settings.claudeKeySecret
			: settings.geminiKeySecret;
	for (const id of [aiKeySecretId(provider), legacy]) {
		if (!id) continue;
		try {
			const key = app.secretStorage.getSecret(id)?.trim();
			if (key) return key;
		} catch {
			// an invalid secret name: try the next one
		}
	}
	return null;
}

function missingKeyMessage(provider: AiProvider): string {
	return `Add a ${AI_PROVIDER_LABELS[provider]} API key in the Free-Excel settings first.`;
}

/** A JSON schema as the tools describe their input (the subset both providers accept). */
export interface JsonSchema {
	type: 'object' | 'string' | 'number' | 'integer' | 'boolean' | 'array';
	description?: string;
	properties?: Record<string, JsonSchema>;
	required?: string[];
	items?: JsonSchema;
	enum?: string[];
}

export interface AiTool {
	name: string;
	description: string;
	parameters: JsonSchema;
}

/** Statuses that mean "busy, try again soon": rate limits and overloaded servers. */
const RETRY_STATUSES = new Set([429, 500, 502, 503, 504, 529]);
/** Waits before each retry of a busy request, in ms (the server's retry-after wins when shorter than the cap). */
const RETRY_DELAYS = [1500, 4000, 9000];
const MAX_RETRY_WAIT = 20000;

/** A request that failed with an HTTP error; `busy` when retrying later may help. */
class AiHttpError extends Error {
	constructor(
		readonly status: number,
		detail: string,
	) {
		super(`${status}: ${detail}`);
	}

	get busy(): boolean {
		return RETRY_STATUSES.has(this.status);
	}
}

function sleep(ms: number): Promise<void> {
	return new Promise((resolve) => window.setTimeout(resolve, ms));
}

function retryAfterMs(headers: Record<string, string>): number | null {
	const value = Object.entries(headers).find(
		([name]) => name.toLowerCase() === 'retry-after',
	)?.[1];
	const seconds = Number(value);
	return value && Number.isFinite(seconds) ? seconds * 1000 : null;
}

async function postJson(
	url: string,
	headers: Record<string, string>,
	body: unknown,
): Promise<unknown> {
	for (let attempt = 0; ; attempt++) {
		const res = await requestUrl({
			url,
			method: 'POST',
			contentType: 'application/json',
			headers,
			body: JSON.stringify(body),
			throw: false,
		});
		let json: unknown = null;
		try {
			json = res.json;
		} catch {
			json = null;
		}
		if (res.status < 400) return json;
		const delay = RETRY_DELAYS[attempt];
		if (RETRY_STATUSES.has(res.status) && delay !== undefined) {
			await sleep(Math.min(retryAfterMs(res.headers) ?? delay, MAX_RETRY_WAIT));
			continue;
		}
		const err = (json as { error?: { message?: unknown } } | null)?.error;
		const detail =
			typeof err?.message === 'string' ? err.message : res.text.slice(0, 300);
		throw new AiHttpError(res.status, detail);
	}
}

/** Explains a failed request in words a user can act on. */
export function aiErrorMessage(provider: AiProvider, err: unknown): string {
	const label = AI_PROVIDER_LABELS[provider];
	if (err instanceof AiHttpError) {
		if (err.status === 429)
			return `${label} rate limit or quota reached (429). Wait a minute and try again, or check the plan of your API key.`;
		if (err.busy)
			return `${label} is overloaded right now (${err.status}) and still was after a few retries. Try again in a minute, or pick another model in the Free-Excel settings.`;
		if (err.status === 401 || err.status === 403)
			return `${label} rejected the API key (${err.status}). Check the key in the Free-Excel settings.`;
	}
	return err instanceof Error ? err.message : String(err);
}

/* ---------------- Claude (Messages API) ---------------- */

interface ClaudeBlock {
	type: string;
	text?: string;
	id?: string;
	name?: string;
	input?: unknown;
}

interface ClaudeResponse {
	content?: ClaudeBlock[];
	stop_reason?: string;
}

async function claudeRequest(
	key: string,
	body: Record<string, unknown>,
): Promise<ClaudeResponse> {
	return (await postJson(
		'https://api.anthropic.com/v1/messages',
		{ 'x-api-key': key, 'anthropic-version': '2023-06-01' },
		body,
	)) as ClaudeResponse;
}

/* ---------------- Gemini (generateContent) ---------------- */

interface GeminiPart {
	text?: string;
	thought?: boolean;
	inlineData?: { mimeType?: string; data?: string };
	functionCall?: { id?: string; name?: string; args?: Record<string, unknown> };
	functionResponse?: unknown;
}

interface GeminiContent {
	role?: string;
	parts?: GeminiPart[];
}

interface GeminiResponse {
	candidates?: { content?: GeminiContent; finishReason?: string }[];
	promptFeedback?: { blockReason?: string };
}

async function geminiRequest(
	key: string,
	model: string,
	body: Record<string, unknown>,
): Promise<GeminiResponse> {
	const id = encodeURIComponent(model.replace(/^models\//, ''));
	return (await postJson(
		`https://generativelanguage.googleapis.com/v1beta/models/${id}:generateContent`,
		{ 'x-goog-api-key': key },
		body,
	)) as GeminiResponse;
}

/** Gemini names the schema types in capitals. */
function geminiSchema(schema: JsonSchema): Record<string, unknown> {
	const out: Record<string, unknown> = { type: schema.type.toUpperCase() };
	if (schema.description) out.description = schema.description;
	if (schema.enum) out.enum = schema.enum;
	if (schema.required) out.required = schema.required;
	if (schema.items) out.items = geminiSchema(schema.items);
	if (schema.properties) {
		const props: Record<string, unknown> = {};
		for (const [name, child] of Object.entries(schema.properties))
			props[name] = geminiSchema(child);
		out.properties = props;
	}
	return out;
}

function geminiText(res: GeminiResponse): string {
	const parts = res.candidates?.[0]?.content?.parts ?? [];
	return parts
		.filter((p) => !p.thought && typeof p.text === 'string')
		.map((p) => p.text)
		.join('');
}

/** Makes a picture with Gemini's image model; throws with a message when none comes back. */
export async function generateGeminiImage(
	app: App,
	settings: PluginSettings,
	prompt: string,
): Promise<{ data: ArrayBuffer; mimeType: string }> {
	const key = aiApiKey(app, settings, 'gemini');
	if (!key)
		throw new Error(
			'Making images needs a Gemini API key: add one in the Free-Excel settings.',
		);
	let res: GeminiResponse;
	try {
		res = await geminiRequest(key, settings.geminiImageModel, {
			contents: [{ role: 'user', parts: [{ text: prompt }] }],
			generationConfig: { responseModalities: ['TEXT', 'IMAGE'] },
		});
	} catch (e) {
		throw new Error(aiErrorMessage('gemini', e));
	}
	if (res.promptFeedback?.blockReason)
		throw new Error(`Blocked: ${res.promptFeedback.blockReason}`);
	const part = (res.candidates?.[0]?.content?.parts ?? []).find(
		(p) => p.inlineData?.data,
	);
	if (!part?.inlineData?.data) {
		const said = geminiText(res).trim();
		throw new Error(
			`Gemini returned no image${said ? `: ${said.slice(0, 300)}` : '.'}`,
		);
	}
	return {
		data: base64ToArrayBuffer(part.inlineData.data),
		mimeType: part.inlineData.mimeType ?? 'image/png',
	};
}

/* ------------------------------------------------------------------ */
/* One question, one answer (AI formulas)                              */
/* ------------------------------------------------------------------ */

const FORMULA_SYSTEM =
	'You are a function inside a spreadsheet cell. Reply with only the value to put in the cell: ' +
	'no explanation, no markdown, no surrounding quotes. Keep the answer short unless the request ' +
	'asks for more. If the request includes cell data, it is given after the request, with tabs ' +
	'between columns and new lines between rows.';

async function askOnce(
	app: App,
	settings: PluginSettings,
	provider: AiProvider,
	prompt: string,
): Promise<string> {
	const key = aiApiKey(app, settings, provider);
	if (!key) throw new Error(missingKeyMessage(provider));
	if (provider === 'claude') {
		const res = await claudeRequest(key, {
			model: settings.claudeModel,
			max_tokens: 1024,
			system: FORMULA_SYSTEM,
			messages: [{ role: 'user', content: prompt }],
		});
		return (res.content ?? [])
			.filter((b) => b.type === 'text')
			.map((b) => b.text ?? '')
			.join('')
			.trim();
	}
	const res = await geminiRequest(key, settings.geminiModel, {
		systemInstruction: { parts: [{ text: FORMULA_SYSTEM }] },
		contents: [{ role: 'user', parts: [{ text: prompt }] }],
	});
	if (res.promptFeedback?.blockReason)
		throw new Error(`Blocked: ${res.promptFeedback.blockReason}`);
	return geminiText(res).trim();
}

/** A short stable hash of a request, used as the key of its cached answer. */
function hashText(text: string): string {
	let h1 = 0xdeadbeef;
	let h2 = 0x41c6ce57;
	for (let i = 0; i < text.length; i++) {
		const ch = text.charCodeAt(i);
		h1 = Math.imul(h1 ^ ch, 2654435761);
		h2 = Math.imul(h2 ^ ch, 1597334677);
	}
	h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
	h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
	return (h2 >>> 0).toString(36) + (h1 >>> 0).toString(36);
}

/**
 * Runs the requests of CLAUDE() and GEMINI() formulas. A formula is evaluated over and over
 * while the sheet is drawn, so every request is sent once: the answer is cached in the file
 * by request, and the cell shows a placeholder until it arrives.
 */
export class AiFormulaRunner {
	private readonly pending = new Set<string>();
	private readonly queue: { id: string; provider: AiProvider; prompt: string }[] =
		[];
	private running = 0;
	/** Requests that failed, by id: not sent again until they are refreshed. */
	private readonly failed = new Map<string, string>();
	private readonly reported = new Set<string>();
	/** Cell ("sheet!row,col") -> the request its formula made last. */
	private readonly cellRequests = new Map<string, string>();

	constructor(
		private readonly app: App,
		private readonly settings: () => PluginSettings,
		/** The file's cached answers (changed in place). */
		private readonly results: () => Record<string, string>,
		/** Called after an answer arrived or a request failed. */
		private readonly onChange: (stored: boolean) => void,
	) {}

	resolve(
		provider: AiProvider,
		prompt: string,
		context: string,
		cell: string | null,
	): Scalar {
		const full = context ? `${prompt}\n\n${context}` : prompt;
		const id = hashText(`${provider}\u0000${full}`);
		if (cell) this.cellRequests.set(cell, id);
		const cached = this.results()[id];
		if (cached !== undefined) return cached;
		if (this.failed.has(id)) return AI_ERROR_TEXT;
		if (!this.pending.has(id)) {
			this.pending.add(id);
			this.queue.push({ id, provider, prompt: full });
			this.pump();
		}
		return AI_LOADING_TEXT;
	}

	/** Forgets the answers of the given cells ("sheet!row,col"), so they are asked again. */
	refreshCells(cells: string[]): number {
		const results = this.results();
		let count = 0;
		for (const cell of cells) {
			const id = this.cellRequests.get(cell);
			if (!id) continue;
			if (id in results || this.failed.has(id)) count++;
			delete results[id];
			this.failed.delete(id);
		}
		return count;
	}

	/** Lets failed requests run again (for example after an API key was added). */
	retryFailed(): void {
		this.failed.clear();
		this.reported.clear();
	}

	private pump(): void {
		while (this.running < MAX_PARALLEL_FORMULAS && this.queue.length > 0) {
			const job = this.queue.shift();
			if (!job) break;
			this.running++;
			void askOnce(this.app, this.settings(), job.provider, job.prompt)
				.then((answer) => {
					const results = this.results();
					results[job.id] = answer;
					const keys = Object.keys(results);
					for (const old of keys.slice(0, Math.max(0, keys.length - MAX_CACHED_RESULTS)))
						delete results[old];
					this.onChange(true);
				})
				.catch((e: unknown) => {
					const message = aiErrorMessage(job.provider, e);
					this.failed.set(job.id, message);
					if (!this.reported.has(message)) {
						this.reported.add(message);
						new Notice(
							`${AI_PROVIDER_LABELS[job.provider]} formula failed: ${message}`,
						);
					}
					this.onChange(false);
				})
				.finally(() => {
					this.pending.delete(job.id);
					this.running--;
					this.pump();
				});
		}
	}
}

/* ------------------------------------------------------------------ */
/* Assistant (chat with tools that read and change the sheet)          */
/* ------------------------------------------------------------------ */

export interface AiToolResult {
	/** What goes back to the model (turned into JSON). */
	result: unknown;
	/** A short line for the chat, such as "Wrote A1:C4". */
	summary?: string;
	isError?: boolean;
}

/** What the assistant needs from the sheet view. */
export interface AiAssistantHost {
	app: App;
	component: Component;
	sourcePath: () => string;
	settings: () => PluginSettings;
	/** The sheet as text, sent with every question. */
	describeSheet: () => string;
	tools: AiTool[];
	runTool: (
		name: string,
		input: Record<string, unknown>,
	) => AiToolResult | Promise<AiToolResult>;
	close: () => void;
}

const ASSISTANT_SYSTEM =
	'You are the assistant of a spreadsheet inside Obsidian (the Free-Excel plugin). You can read ' +
	'the sheet and change it with the tools. Cell references use A1 notation; rows start at 1. ' +
	'Formulas start with "=" and use Excel syntax (SUM, IF, VLOOKUP, COUNTIF, ...). Prefer formulas ' +
	'over computed numbers when the user may change the data later. Before changing cells you have ' +
	'not seen, read them. Do not overwrite cells that hold data unless the user asked for it. ' +
	'Every change can be undone by the user with Ctrl+Z. After making changes, say briefly what ' +
	'you did. Images can be put in a cell or float over the sheet: find_images looks for pictures ' +
	'in the vault, search_web_images looks for pictures on the internet (Wikimedia Commons), ' +
	'insert_image adds one from the vault or a web address, and generate_image makes a new ' +
	'picture. Never guess image addresses: use only addresses the user gave or search_web_images ' +
	'returned, and pick one that is not already in the sheet. An emoji in a cell is not an image. Answer in the language the user writes in. ' +
	'Keep answers short and use Markdown.';

type ChatEntry =
	| { kind: 'user'; text: string }
	| { kind: 'assistant'; text: string }
	| { kind: 'action'; text: string; isError: boolean }
	| { kind: 'error'; text: string };

/** The chat panel: a conversation with Claude or Gemini that can read and change the sheet. */
export class AiAssistantPanel {
	private provider: AiProvider;
	private claudeHistory: unknown[] = [];
	private geminiHistory: GeminiContent[] = [];
	private entries: ChatEntry[] = [];
	private busy = false;
	private readonly listEl: HTMLElement;
	private readonly inputEl: HTMLTextAreaElement;
	private readonly sendBtn: HTMLButtonElement;
	private readonly providerSelect: HTMLSelectElement;
	private readonly renderChild = new Component();
	/** Changes while a reply is running when the chat is cleared: the old reply is dropped. */
	private generation = 0;

	constructor(
		private readonly el: HTMLElement,
		private readonly host: AiAssistantHost,
	) {
		this.provider = host.settings().aiProvider;
		host.component.addChild(this.renderChild);
		el.empty();
		el.addClass('rtl-sheet-ai-panel');

		const head = el.createDiv({ cls: 'rtl-sheet-side-head' });
		const title = head.createDiv({ cls: 'rtl-sheet-side-title' });
		setIcon(title.createSpan(), 'sparkles');
		title.createSpan({ text: 'AI assistant' });
		this.providerSelect = head.createEl('select', {
			cls: 'dropdown rtl-sheet-ai-provider',
			attr: { 'aria-label': 'AI provider' },
		});
		for (const p of ['claude', 'gemini'] as const)
			this.providerSelect.createEl('option', {
				value: p,
				text: AI_PROVIDER_LABELS[p],
			});
		this.providerSelect.value = this.provider;
		this.providerSelect.addEventListener('change', () => {
			this.provider = this.providerSelect.value === 'gemini' ? 'gemini' : 'claude';
			this.newChat();
		});
		this.iconButton(head, 'rotate-ccw', 'New chat', () => this.newChat());
		this.iconButton(head, 'x', 'Close', () => host.close());

		this.listEl = el.createDiv({ cls: 'rtl-sheet-ai-list' });

		const form = el.createDiv({ cls: 'rtl-sheet-ai-form' });
		this.inputEl = form.createEl('textarea', {
			cls: 'rtl-sheet-ai-input',
			attr: {
				placeholder: 'Ask about the sheet, or ask for a change…',
				dir: 'auto',
				rows: '3',
			},
		});
		this.inputEl.addEventListener('keydown', (e) => {
			if (e.key === 'Enter' && !e.shiftKey && !e.isComposing) {
				e.preventDefault();
				void this.send();
			}
		});
		this.sendBtn = this.iconButton(form, 'send', 'Send', () => void this.send());
		this.sendBtn.addClass('mod-cta');
		this.render();
	}

	focus(): void {
		this.inputEl.focus();
	}

	destroy(): void {
		this.generation++;
		this.host.component.removeChild(this.renderChild);
		this.el.empty();
		this.el.removeClass('rtl-sheet-ai-panel');
	}

	private iconButton(
		parent: HTMLElement,
		icon: string,
		label: string,
		onClick: () => void,
	): HTMLButtonElement {
		const btn = parent.createEl('button', {
			cls: 'rtl-sheet-side-btn',
			attr: { 'aria-label': label, title: label },
		});
		setIcon(btn, icon);
		btn.addEventListener('click', onClick);
		return btn;
	}

	private newChat(): void {
		this.generation++;
		this.busy = false;
		this.claudeHistory = [];
		this.geminiHistory = [];
		this.entries = [];
		this.render();
	}

	private render(): void {
		const list = this.listEl;
		list.empty();
		this.renderChild.unload();
		this.renderChild.load();
		if (this.entries.length === 0) {
			const hint = list.createDiv({ cls: 'rtl-sheet-ai-hint' });
			hint.createEl('p', {
				text: `${AI_PROVIDER_LABELS[this.provider]} sees this sheet and can change it. Every change can be undone with Ctrl+Z.`,
			});
			hint.createEl('p', {
				text: 'For example: add a total row, highlight the values above 100, explain a formula, or fill a column from the data next to it.',
			});
			if (!aiApiKey(this.host.app, this.host.settings(), this.provider))
				hint.createEl('p', {
					cls: 'rtl-sheet-ai-warning',
					text: missingKeyMessage(this.provider),
				});
		}
		for (const entry of this.entries) {
			const row = list.createDiv({ cls: `rtl-sheet-ai-msg is-${entry.kind}` });
			if (entry.kind === 'assistant') {
				row.setAttr('dir', 'auto');
				void MarkdownRenderer.render(
					this.host.app,
					entry.text,
					row,
					this.host.sourcePath(),
					this.renderChild,
				);
			} else if (entry.kind === 'action') {
				setIcon(row.createSpan(), entry.isError ? 'alert-triangle' : 'check');
				row.createSpan({ text: entry.text });
			} else {
				row.setText(entry.text);
				row.setAttr('dir', 'auto');
			}
		}
		if (this.busy)
			list.createDiv({
				cls: 'rtl-sheet-ai-msg is-busy',
				text: `${AI_PROVIDER_LABELS[this.provider]} is working…`,
			});
		this.sendBtn.disabled = this.busy;
		list.scrollTop = list.scrollHeight;
	}

	private add(entry: ChatEntry): void {
		this.entries.push(entry);
		this.render();
	}

	private async send(): Promise<void> {
		const text = this.inputEl.value.trim();
		if (!text || this.busy) return;
		const key = aiApiKey(this.host.app, this.host.settings(), this.provider);
		if (!key) {
			new Notice(missingKeyMessage(this.provider));
			return;
		}
		this.inputEl.value = '';
		this.busy = true;
		this.add({ kind: 'user', text });
		const generation = this.generation;
		// The sheet as it is now goes with the question, so the model does not have to ask for it.
		const message = `${text}\n\n<sheet>\n${this.host.describeSheet()}\n</sheet>`;
		try {
			if (this.provider === 'claude')
				await this.runClaude(key, message, generation);
			else await this.runGemini(key, message, generation);
		} catch (e) {
			if (generation === this.generation)
				this.add({
					kind: 'error',
					text: aiErrorMessage(this.provider, e),
				});
		} finally {
			if (generation === this.generation) {
				this.busy = false;
				this.render();
			}
		}
	}

	private async runTool(
		name: string,
		input: Record<string, unknown>,
	): Promise<{ result: unknown; isError: boolean }> {
		let out: AiToolResult;
		try {
			out = await this.host.runTool(name, input);
		} catch (e) {
			const error = e instanceof Error ? e.message : String(e);
			out = { result: { error }, summary: `${name} failed: ${error}`, isError: true };
		}
		if (out.summary || out.isError)
			this.add({
				kind: 'action',
				text: out.summary ?? `${name} failed`,
				isError: !!out.isError,
			});
		return { result: out.result, isError: !!out.isError };
	}

	private async runClaude(
		key: string,
		message: string,
		generation: number,
	): Promise<void> {
		const history = this.claudeHistory;
		history.push({ role: 'user', content: message });
		const tools = this.host.tools.map((t) => ({
			name: t.name,
			description: t.description,
			input_schema: t.parameters,
		}));
		for (let step = 0; step < MAX_AGENT_STEPS; step++) {
			const res = await claudeRequest(key, {
				model: this.host.settings().claudeModel,
				max_tokens: 8192,
				system: ASSISTANT_SYSTEM,
				tools,
				messages: history,
			});
			if (generation !== this.generation) return;
			const content = res.content ?? [];
			history.push({ role: 'assistant', content });
			const text = content
				.filter((b) => b.type === 'text')
				.map((b) => b.text ?? '')
				.join('')
				.trim();
			if (text) this.add({ kind: 'assistant', text });
			const calls = content.filter((b) => b.type === 'tool_use');
			if (res.stop_reason !== 'tool_use' || calls.length === 0) return;
			// One after another: a tool may depend on what the one before it changed.
			const results: unknown[] = [];
			for (const call of calls) {
				const { result, isError } = await this.runTool(
					call.name ?? '',
					(call.input ?? {}) as Record<string, unknown>,
				);
				results.push({
					type: 'tool_result',
					tool_use_id: call.id,
					content: JSON.stringify(result),
					is_error: isError,
				});
			}
			if (generation !== this.generation) return;
			history.push({ role: 'user', content: results });
		}
		this.add({ kind: 'error', text: 'Stopped after too many steps.' });
	}

	private async runGemini(
		key: string,
		message: string,
		generation: number,
	): Promise<void> {
		const history = this.geminiHistory;
		history.push({ role: 'user', parts: [{ text: message }] });
		const tools = [
			{
				functionDeclarations: this.host.tools.map((t) => ({
					name: t.name,
					description: t.description,
					parameters: geminiSchema(t.parameters),
				})),
			},
		];
		for (let step = 0; step < MAX_AGENT_STEPS; step++) {
			const res = await geminiRequest(key, this.host.settings().geminiModel, {
				systemInstruction: { parts: [{ text: ASSISTANT_SYSTEM }] },
				contents: history,
				tools,
			});
			if (generation !== this.generation) return;
			if (res.promptFeedback?.blockReason)
				throw new Error(`Blocked: ${res.promptFeedback.blockReason}`);
			const content = res.candidates?.[0]?.content;
			if (!content?.parts) {
				const reason = res.candidates?.[0]?.finishReason;
				throw new Error(`No answer${reason ? ` (${reason})` : ''}.`);
			}
			// Sent back as it came: Gemini needs the signatures of its own function calls.
			history.push({ role: 'model', parts: content.parts });
			const text = geminiText(res).trim();
			if (text) this.add({ kind: 'assistant', text });
			const calls = content.parts.filter((p) => p.functionCall);
			if (calls.length === 0) return;
			const parts: GeminiPart[] = [];
			for (const p of calls) {
				const call = p.functionCall ?? {};
				const { result, isError } = await this.runTool(call.name ?? '', call.args ?? {});
				parts.push({
					functionResponse: {
						...(call.id ? { id: call.id } : {}),
						name: call.name,
						response: isError ? { error: result } : { result },
					},
				});
			}
			if (generation !== this.generation) return;
			history.push({ role: 'user', parts });
		}
		this.add({ kind: 'error', text: 'Stopped after too many steps.' });
	}
}
