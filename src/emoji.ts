/** One emoji and the words it can be found by. */
export interface EmojiInfo {
	char: string;
	keywords: string;
}

export interface EmojiCategory {
	name: string;
	icon: string;
	emojis: EmojiInfo[];
}

/** "😀 grinning smile|😃 ..." -> the emojis of a category. */
function parse(list: string): EmojiInfo[] {
	return list
		.split('|')
		.map((item) => item.trim())
		.filter((item) => item !== '')
		.map((item) => {
			const space = item.indexOf(' ');
			return {
				char: space < 0 ? item : item.slice(0, space),
				keywords: space < 0 ? '' : item.slice(space + 1).toLowerCase(),
			};
		});
}

export const EMOJI_CATEGORIES: EmojiCategory[] = [
	{
		name: 'Smileys',
		icon: 'smile',
		emojis: parse(
			'😀 grinning smile happy|😃 smiley happy joy|😄 smile happy laugh|😁 grin beaming|😆 laughing satisfied|😅 sweat smile relief|🤣 rolling floor laughing rofl|😂 joy tears laughing lol|🙂 slightly smiling|🙃 upside down|😉 wink|😊 blush smiling|😇 innocent halo angel|🥰 love hearts adore|😍 heart eyes love|🤩 star struck excited|😘 kiss blow|😗 kissing|😚 kissing closed eyes|😋 yum delicious tongue|😛 tongue|😜 wink tongue crazy|🤪 zany crazy|😝 squinting tongue|🤑 money mouth rich|🤗 hug hugging|🤭 hand over mouth oops|🤫 shush quiet secret|🤔 thinking hmm|🤐 zipper mouth|🤨 raised eyebrow skeptical|😐 neutral|😑 expressionless|😶 no mouth silent|😏 smirk|😒 unamused|🙄 eye roll|😬 grimace awkward|😌 relieved calm|😔 pensive sad|😪 sleepy|🤤 drooling|😴 sleeping zzz|😷 mask sick|🤒 thermometer sick fever|🤕 bandage hurt|🤢 nauseated sick|🤮 vomit|🥵 hot sweating|🥶 cold freezing|🥴 woozy dizzy|😵 dizzy|🤯 exploding head mind blown|🤠 cowboy|🥳 party celebrate birthday|😎 cool sunglasses|🤓 nerd geek|🧐 monocle inspect|😕 confused|😟 worried|🙁 frowning sad|😮 open mouth surprised|😯 hushed|😲 astonished shocked|😳 flushed embarrassed|🥺 pleading puppy eyes|😦 frowning open mouth|😧 anguished|😨 fearful scared|😰 anxious sweat|😥 sad relieved|😢 crying sad tear|😭 sobbing loudly crying|😱 scream fear|😖 confounded|😣 persevering|😞 disappointed|😓 downcast sweat|😩 weary tired|😫 tired exhausted|🥱 yawn bored|😤 triumph steam angry|😡 pouting angry rage|😠 angry mad|🤬 cursing swearing|😈 devil smiling|👿 imp angry devil|💀 skull dead|☠️ skull crossbones danger|💩 poop|🤡 clown|👻 ghost|👽 alien|🤖 robot bot|😺 cat smiling|😸 cat grin|😹 cat joy|😻 cat heart eyes|🙈 see no evil monkey|🙉 hear no evil monkey|🙊 speak no evil monkey',
		),
	},
	{
		name: 'People',
		icon: 'hand',
		emojis: parse(
			'👋 wave hello bye|🤚 raised back hand|✋ raised hand stop high five|🖖 vulcan salute|👌 ok perfect|🤌 pinched fingers|🤏 pinching small|✌️ victory peace|🤞 crossed fingers luck|🤟 love you gesture|🤘 rock horns|🤙 call me|👈 point left|👉 point right|👆 point up|👇 point down|☝️ index up|👍 thumbs up like yes approve|👎 thumbs down dislike no|✊ raised fist|👊 fist bump punch|👏 clap applause|🙌 raising hands celebrate|👐 open hands|🤲 palms up|🤝 handshake deal agreement|🙏 pray please thanks|✍️ writing|💅 nail polish|💪 muscle strong flex|🧠 brain smart|👀 eyes look see|👁️ eye|👶 baby|🧒 child kid|👦 boy|👧 girl|🧑 person adult|👨 man|👩 woman|🧓 older person|👴 old man grandpa|👵 old woman grandma|🙋 raising hand|🙅 no gesture|🙆 ok gesture|💁 tipping hand info|🤷 shrug dunno|🤦 facepalm|🧑‍💻 technologist developer coder|🧑‍🏫 teacher|🧑‍🎓 student graduate|🧑‍⚕️ health worker doctor|🧑‍🍳 cook chef|🧑‍🔧 mechanic|🧑‍🔬 scientist|🧑‍🎨 artist|🧑‍🚀 astronaut|👮 police officer|👷 construction worker|💂 guard|🕵️ detective|👸 princess|🤴 prince|🦸 superhero|🧙 mage wizard|🏃 running|🚶 walking|💃 dancing woman|🕺 dancing man|👯 bunny ears party|🧘 yoga meditate|👪 family|👫 couple|💑 couple heart|🗣️ speaking head|👤 bust silhouette user|👥 users people group',
		),
	},
	{
		name: 'Hearts and symbols',
		icon: 'heart',
		emojis: parse(
			'❤️ red heart love|🧡 orange heart|💛 yellow heart|💚 green heart|💙 blue heart|💜 purple heart|🖤 black heart|🤍 white heart|🤎 brown heart|💔 broken heart|❣️ heart exclamation|💕 two hearts|💞 revolving hearts|💓 beating heart|💗 growing heart|💖 sparkling heart|💘 cupid arrow heart|💝 heart gift ribbon|✅ check mark done yes complete|☑️ ballot check box|✔️ check mark|❌ cross mark no wrong cancel|❎ cross button|➕ plus add|➖ minus subtract|➗ divide|✖️ multiply|❓ question mark|❔ white question|❗ exclamation important|❕ white exclamation|‼️ double exclamation|⁉️ exclamation question|⚠️ warning caution|🚫 prohibited forbidden no|⛔ no entry|🔴 red circle|🟠 orange circle|🟡 yellow circle|🟢 green circle|🔵 blue circle|🟣 purple circle|⚫ black circle|⚪ white circle|🟤 brown circle|🟥 red square|🟧 orange square|🟨 yellow square|🟩 green square|🟦 blue square|🟪 purple square|⬛ black square|⬜ white square|🔶 orange diamond|🔷 blue diamond|🔺 red triangle up|🔻 red triangle down|⭐ star|🌟 glowing star|✨ sparkles|⚡ lightning bolt high voltage|🔥 fire hot lit|💯 hundred perfect score|💢 anger|💥 collision boom|💫 dizzy|💬 speech balloon comment|💭 thought balloon|🗯️ anger bubble|💤 zzz sleep|♻️ recycle|✳️ asterisk|❇️ sparkle|🆗 ok button|🆕 new|🆓 free|🆒 cool|🆙 up|🔝 top|🔜 soon|🔙 back|🔛 on|🔚 end|ℹ️ information|Ⓜ️ m circled|🅰️ a blood type|🅱️ b blood type|©️ copyright|®️ registered|™️ trade mark|#️⃣ hash number|0️⃣ zero|1️⃣ one|2️⃣ two|3️⃣ three|4️⃣ four|5️⃣ five|6️⃣ six|7️⃣ seven|8️⃣ eight|9️⃣ nine|🔟 ten|⬆️ up arrow|⬇️ down arrow|⬅️ left arrow|➡️ right arrow|↗️ up right arrow|↘️ down right arrow|↙️ down left arrow|↖️ up left arrow|↕️ up down arrow|↔️ left right arrow|🔄 counterclockwise refresh|🔃 clockwise|▶️ play|⏸️ pause|⏹️ stop|⏺️ record|⏩ fast forward|⏪ rewind|🔀 shuffle|🔁 repeat|➰ curly loop|〰️ wavy dash|💲 dollar sign|💱 currency exchange|☮️ peace|☯️ yin yang|✡️ star of david|🕎 menorah|☪️ star crescent|✝️ cross|🔯 six pointed star|♈ aries|♉ taurus|♊ gemini|♋ cancer|♌ leo|♍ virgo|♎ libra|♏ scorpio|♐ sagittarius|♑ capricorn|♒ aquarius|♓ pisces',
		),
	},
	{
		name: 'Animals and nature',
		icon: 'leaf',
		emojis: parse(
			'🐶 dog puppy|🐱 cat kitten|🐭 mouse|🐹 hamster|🐰 rabbit bunny|🦊 fox|🐻 bear|🐼 panda|🐨 koala|🐯 tiger|🦁 lion|🐮 cow|🐷 pig|🐸 frog|🐵 monkey|🐔 chicken|🐧 penguin|🐦 bird|🐤 chick|🦆 duck|🦅 eagle|🦉 owl|🦇 bat|🐺 wolf|🐴 horse|🦄 unicorn|🐝 bee honeybee|🐛 bug caterpillar|🦋 butterfly|🐌 snail slow|🐞 ladybug|🐜 ant|🕷️ spider|🐢 turtle slow|🐍 snake|🦎 lizard|🐙 octopus|🦑 squid|🦀 crab|🐠 tropical fish|🐟 fish|🐬 dolphin|🐳 whale|🦈 shark|🐊 crocodile|🐘 elephant|🦒 giraffe|🦓 zebra|🐪 camel|🐄 cow|🐑 sheep|🐐 goat|🐓 rooster|🦃 turkey|🕊️ dove peace|🐕 dog|🐈 cat|🌵 cactus|🎄 christmas tree|🌲 evergreen tree|🌳 tree|🌴 palm tree|🌱 seedling sprout grow|🌿 herb|☘️ shamrock|🍀 four leaf clover luck|🍁 maple leaf|🍂 fallen leaves autumn|🍃 leaves wind|🌾 rice|🌷 tulip flower|🌹 rose flower|🥀 wilted flower|🌺 hibiscus|🌸 cherry blossom|🌼 blossom|🌻 sunflower|💐 bouquet flowers|🍄 mushroom|🌍 earth globe europe africa|🌎 earth globe americas|🌏 earth globe asia|🌕 full moon|🌙 crescent moon|☀️ sun sunny|🌤️ sun small cloud|⛅ partly cloudy|☁️ cloud|🌧️ rain|⛈️ thunderstorm|🌩️ lightning|❄️ snowflake cold|☃️ snowman|🌈 rainbow|☔ umbrella rain|💧 droplet water|🌊 wave ocean|🌪️ tornado|🌫️ fog',
		),
	},
	{
		name: 'Food and drink',
		icon: 'apple',
		emojis: parse(
			'🍏 green apple|🍎 red apple|🍐 pear|🍊 orange tangerine|🍋 lemon|🍌 banana|🍉 watermelon|🍇 grapes|🍓 strawberry|🫐 blueberries|🍈 melon|🍒 cherries|🍑 peach|🥭 mango|🍍 pineapple|🥥 coconut|🥝 kiwi|🍅 tomato|🍆 eggplant|🥑 avocado|🥦 broccoli|🥬 leafy green|🥒 cucumber|🌶️ hot pepper chili|🌽 corn|🥕 carrot|🧄 garlic|🧅 onion|🥔 potato|🥐 croissant|🥯 bagel|🍞 bread|🥖 baguette|🧀 cheese|🥚 egg|🍳 cooking fried egg|🥞 pancakes|🧇 waffle|🥓 bacon|🍗 poultry leg chicken|🍖 meat bone|🌭 hot dog|🍔 hamburger burger|🍟 fries|🍕 pizza|🥪 sandwich|🌮 taco|🌯 burrito|🥗 salad|🍝 spaghetti pasta|🍜 ramen noodles|🍲 stew pot|🍛 curry|🍣 sushi|🍱 bento|🥟 dumpling|🍤 shrimp|🍚 rice|🍙 rice ball|🍦 ice cream soft|🍨 ice cream|🍩 doughnut donut|🍪 cookie|🎂 birthday cake|🍰 cake shortcake|🧁 cupcake|🥧 pie|🍫 chocolate|🍬 candy|🍭 lollipop|🍯 honey|☕ coffee hot beverage|🍵 tea|🧃 juice box|🥤 cup straw soda|🧋 bubble tea|🍺 beer|🍻 cheers beers|🥂 clinking glasses toast|🍷 wine|🥃 whisky tumbler|🍸 cocktail|🍹 tropical drink|🍾 champagne bottle|🧊 ice|🥛 milk|🍽️ plate fork knife|🍴 fork knife',
		),
	},
	{
		name: 'Activities and travel',
		icon: 'plane',
		emojis: parse(
			'⚽ soccer football|🏀 basketball|🏈 american football|⚾ baseball|🎾 tennis|🏐 volleyball|🏉 rugby|🎱 billiards 8 ball|🏓 ping pong|🏸 badminton|🥅 goal net|⛳ golf|🏹 bow arrow|🎣 fishing|🥊 boxing|🥋 martial arts|⛸️ ice skate|🎿 ski|🏂 snowboard|🏋️ weight lifting gym|🚴 cycling bike|🏊 swimming|🏆 trophy winner|🥇 gold medal first|🥈 silver medal second|🥉 bronze medal third|🏅 medal|🎖️ military medal|🎗️ ribbon|🎫 ticket|🎟️ admission tickets|🎪 circus|🎭 theater|🎨 art palette|🎬 movie clapper|🎤 microphone sing|🎧 headphones music|🎼 score music|🎹 piano keyboard|🥁 drum|🎷 saxophone|🎺 trumpet|🎸 guitar|🎻 violin|🎲 dice game|♟️ chess|🎯 target bullseye goal|🎳 bowling|🎮 video game controller|🧩 puzzle|🎉 party popper celebrate tada|🎊 confetti|🎈 balloon|🎁 gift present|🎀 ribbon bow|🎃 jack o lantern halloween|🎄 christmas|🎆 fireworks|🎇 sparkler|🧨 firecracker|🚗 car|🚕 taxi|🚌 bus|🚎 trolleybus|🏎️ racing car|🚓 police car|🚑 ambulance|🚒 fire engine|🚚 truck delivery|🚜 tractor|🛵 scooter|🏍️ motorcycle|🚲 bicycle|🛴 kick scooter|🚂 train locomotive|🚆 train|🚇 metro subway|✈️ airplane flight|🛫 departure|🛬 arrival|🚀 rocket launch|🛸 ufo|🚁 helicopter|⛵ sailboat|🚤 speedboat|🛳️ ship cruise|⚓ anchor|⛽ fuel|🚦 traffic light|🚧 construction|🗺️ map|🧭 compass|🏔️ mountain snow|⛰️ mountain|🏕️ camping|🏖️ beach|🏝️ island|🏠 house home|🏡 house garden|🏢 office building|🏥 hospital|🏦 bank|🏨 hotel|🏪 store shop|🏫 school|🏛️ classical building|⛪ church|🕌 mosque|🕍 synagogue|🗽 statue of liberty|🗼 tokyo tower|🏰 castle|🎡 ferris wheel|🎢 roller coaster|🌅 sunrise|🌄 sunrise mountains|🌃 night city|🌉 bridge night',
		),
	},
	{
		name: 'Objects',
		icon: 'lightbulb',
		emojis: parse(
			'⌚ watch|📱 mobile phone|💻 laptop computer|⌨️ keyboard|🖥️ desktop computer|🖨️ printer|🖱️ mouse computer|💾 floppy disk save|💿 cd disc|📷 camera|📹 video camera|🎥 movie camera|📞 telephone receiver|☎️ telephone|📺 television tv|📻 radio|🎙️ studio microphone|⏰ alarm clock|⏳ hourglass time|⌛ hourglass done|📡 satellite antenna|🔋 battery|🔌 plug electric|💡 light bulb idea|🔦 flashlight|🕯️ candle|🧯 fire extinguisher|💸 money wings|💵 dollar banknote|💶 euro banknote|💷 pound banknote|💴 yen banknote|💰 money bag|💳 credit card|🧾 receipt|💎 gem diamond|⚖️ balance scale law|🔧 wrench tool|🔨 hammer|⚒️ hammer pick|🛠️ tools|⛏️ pick|🔩 nut bolt|⚙️ gear settings|🧱 brick|⛓️ chains|🧲 magnet|🔫 water pistol|💣 bomb|🔪 knife|🛡️ shield|🔮 crystal ball|🧿 nazar amulet|💈 barber|⚗️ alembic|🔭 telescope|🔬 microscope|🩺 stethoscope|💊 pill medicine|💉 syringe vaccine|🩹 bandage|🧬 dna|🦠 microbe virus|🧪 test tube|🌡️ thermometer|🧹 broom|🧺 basket|🧻 toilet paper|🚽 toilet|🚿 shower|🛁 bathtub|🧼 soap|🧽 sponge|🔑 key|🗝️ old key|🚪 door|🪑 chair|🛋️ couch|🛏️ bed|🧸 teddy bear|🖼️ framed picture|🛍️ shopping bags|🛒 shopping cart|🎁 gift|✉️ envelope email|📧 e-mail|📨 incoming envelope|📩 envelope arrow|📤 outbox|📥 inbox|📦 package box|🏷️ label tag|📪 mailbox closed|📫 mailbox|📮 postbox|📜 scroll|📃 page curl|📄 page document|📑 bookmark tabs|📊 bar chart|📈 chart increasing up trend|📉 chart decreasing down trend|🗒️ notepad|🗓️ calendar spiral|📆 calendar tear off|📅 calendar date|🗑️ wastebasket trash delete|📇 card index|🗃️ card file box|🗳️ ballot box|🗄️ file cabinet|📋 clipboard|📁 folder|📂 open folder|🗂️ dividers|🗞️ newspaper rolled|📰 newspaper|📓 notebook|📔 notebook decorative|📒 ledger|📕 closed book red|📗 green book|📘 blue book|📙 orange book|📚 books library|📖 open book read|🔖 bookmark|🧷 safety pin|🔗 link chain|📎 paperclip attachment|🖇️ linked paperclips|📐 triangular ruler|📏 ruler|🧮 abacus|📌 pushpin pin|📍 round pushpin location|✂️ scissors cut|🖊️ pen|🖋️ fountain pen|✒️ nib|🖌️ paintbrush|🖍️ crayon|📝 memo note write|✏️ pencil edit|🔍 magnifying glass search|🔎 magnifying glass right|🔏 locked pen|🔐 locked key|🔒 locked lock secure|🔓 unlocked|🔔 bell notification|🔕 bell off mute|📣 megaphone|📢 loudspeaker announcement|🏁 chequered flag finish|🚩 triangular flag|🎌 crossed flags|🏴 black flag|🏳️ white flag|🏳️‍🌈 rainbow flag pride|🇮🇱 israel flag|🇺🇸 united states usa flag|🇬🇧 united kingdom uk flag|🇪🇺 european union flag|🇩🇪 germany flag|🇫🇷 france flag|🇪🇸 spain flag|🇮🇹 italy flag|🇯🇵 japan flag|🇨🇳 china flag|🇮🇳 india flag|🇧🇷 brazil flag|🇨🇦 canada flag|🇺🇦 ukraine flag',
		),
	},
];

/** Every emoji once, for searching. */
export const ALL_EMOJIS: EmojiInfo[] = (() => {
	const seen = new Set<string>();
	const out: EmojiInfo[] = [];
	for (const category of EMOJI_CATEGORIES)
		for (const e of category.emojis)
			if (!seen.has(e.char)) {
				seen.add(e.char);
				out.push(e);
			}
	return out;
})();

/** Emojis whose words start with (first) or contain the query. */
export function searchEmojis(query: string): EmojiInfo[] {
	const q = query.trim().toLowerCase();
	if (!q) return ALL_EMOJIS;
	const words = q.split(/\s+/);
	const scored: { e: EmojiInfo; score: number }[] = [];
	for (const e of ALL_EMOJIS) {
		const keywords = e.keywords.split(' ');
		let score = 0;
		let all = true;
		for (const w of words) {
			if (keywords.some((k) => k === w)) score += 3;
			else if (keywords.some((k) => k.startsWith(w))) score += 2;
			else if (e.keywords.includes(w)) score += 1;
			else {
				all = false;
				break;
			}
		}
		if (all) scored.push({ e, score });
	}
	return scored.sort((a, b) => b.score - a.score).map((s) => s.e);
}
