/* global chrome */
(function () {
  'use strict';

  // ── Detection ────────────────────────────────────────────────────────────────

  const MD_EXTS = ['.md', '.markdown', '.mdown', '.mkd', '.mkdn', '.mdwn', '.mdtext'];

  function isMarkdownPage() {
    const url = location.href.toLowerCase().split('?')[0].split('#')[0];
    if (!MD_EXTS.some(e => url.endsWith(e))) return false;

    // Browsers display plain-text files as a bare <pre> inside <body>
    const body = document.body;
    if (!body) return false;
    const realChildren = [...body.children].filter(
      el => !['SCRIPT', 'STYLE', 'LINK', 'META'].includes(el.tagName)
    );
    return realChildren.length === 0 ||
      (realChildren.length === 1 && realChildren[0].tagName === 'PRE');
  }

  // ── Utilities ────────────────────────────────────────────────────────────────

  function esc(s) {
    return String(s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function slugify(s) {
    return s.replace(/<[^>]+>/g, '').toLowerCase()
      .replace(/[^\w\s-]/g, '').trim().replace(/[\s]+/g, '-');
  }

  function getFilename() {
    return decodeURIComponent(location.pathname.split('/').pop() || 'document.md');
  }

  // ── Syntax Highlighting ──────────────────────────────────────────────────────

  const LANG_ALIASES = {
    js: 'javascript', jsx: 'javascript', mjs: 'javascript', cjs: 'javascript',
    ts: 'typescript', tsx: 'typescript',
    py: 'python', py3: 'python',
    sh: 'bash', shell: 'bash', zsh: 'bash',
    rb: 'ruby', yml: 'yaml', md: 'markdown',
    rest: 'http', graphql: 'http'
  };

  const KEYWORDS = {
    javascript: 'async await break case catch class const continue debugger default delete do else export extends finally for from function if import in instanceof let new null of return static super switch this throw true false try typeof undefined var void while with yield',
    typescript: 'abstract any async await break case catch class const continue declare default delete do else enum export extends finally for from function if implements import in instanceof interface keyof let namespace new null of override package private protected public readonly return static super switch this throw true false try type typeof undefined var void while with yield',
    python: 'and as assert async await break class continue def del elif else except False finally for from global if import in is lambda None nonlocal not or pass raise return True try while with yield',
    java: 'abstract assert boolean break byte case catch char class const continue default do double else enum extends final finally float for goto if implements import instanceof int interface long native new null package private protected public return short static strictfp super switch synchronized this throw throws transient true false try var void volatile while',
    go: 'break case chan const continue default defer else fallthrough false for func go goto if import interface map nil package range return select struct switch true type var',
    rust: 'as async await break const continue crate dyn else enum extern false fn for if impl in let loop match mod move mut pub ref return self Self static struct super trait true type union unsafe use where while',
    csharp: 'abstract as base bool break byte case catch char checked class const continue decimal default delegate do double else enum event explicit extern false finally fixed float for foreach goto if implicit in int interface internal is lock long namespace new null object operator out override params private protected public readonly ref return sbyte sealed short sizeof stackalloc static string struct switch this throw true try typeof uint ulong unchecked unsafe ushort using virtual void volatile while',
    php: 'abstract and array as break callable case catch class clone const continue declare default do echo else elseif empty enddeclare endfor endforeach endif endswitch endwhile extends false final finally fn for foreach function global goto if implements include include_once instanceof insteadof interface list match namespace new null or print private protected public readonly require require_once return static switch throw trait true try use var while xor yield',
    ruby: 'alias and begin break case class def defined? do else elsif end ensure false for if in module next nil not or redo rescue retry return self super then true undef unless until when while yield',
    swift: 'Any as associatedtype associativity break case catch class continue convenience deinit default defer do dynamic else enum extension false fallthrough fileprivate final for func get guard if import in infix init inout internal is lazy left let mutating nil none nonmutating open operator optional override postfix precedence prefix private protocol public repeat required rethrows return right self Self set static struct subscript super switch throw throws true try type typealias unowned var weak where while',
    kotlin: 'abstract actual as break by catch class companion const constructor continue crossinline data delegate do dynamic else enum expect external false field file finally for fun get if in infix init inline inner interface internal is it lateinit noinline null object open operator out override package param private protected public receiver reified return sealed set setparam super suspend tailrec this throw true try typealias typeof val var vararg when where while',
    sql: 'ADD ALL ALTER AND AS ASC BETWEEN BY CASE COLUMN CONSTRAINT CREATE CROSS DATABASE DEFAULT DELETE DESC DISTINCT DROP ELSE END EXISTS EXPLAIN FALSE FOREIGN FROM FULL GROUP HAVING IF IN INDEX INNER INSERT INTO IS JOIN KEY LEFT LIKE LIMIT NOT NULL ON OR ORDER OUTER PRIMARY RIGHT SELECT SET TABLE THEN TRUE TRUNCATE UNION UNIQUE UPDATE VALUES VIEW WHERE WITH',
    bash: 'case do done elif else esac fi for function if in select then time until while',
    css: '',
    html: '',
    json: '',
    xml: '',
    yaml: '',
    markdown: '',
    http: ''
  };

  function highlight(code, lang) {
    const normalized = LANG_ALIASES[lang] || lang;
    if (!(normalized in KEYWORDS)) {
      return code;
    }

    let result = code;
    const placeholder = [];

    // Protect already-escaped HTML entities
    result = result.replace(/(&(?:[a-z]+|#\d+);)/g, (m) => {
      placeholder.push(m);
      return `\x00P${placeholder.length - 1}\x00`;
    });

    const addSpan = (cls, text) => `<span class="sh-${cls}">${text}</span>`;

    // Strings
    result = result.replace(/(["'`])(?:\\[\s\S]|(?!\1)[^\\])*\1/g, m => addSpan('str', m));

    // Comments
    if (['javascript','typescript','java','go','rust','csharp','swift','kotlin','php','css'].includes(normalized)) {
      result = result.replace(/\/\/[^\n]*/g, m => addSpan('cmt', m));
      result = result.replace(/\/\*[\s\S]*?\*\//g, m => addSpan('cmt', m));
    }
    if (['python','ruby','bash','yaml'].includes(normalized)) {
      result = result.replace(/(^|\n)([ \t]*)#[^\n]*/g, (m, nl, indent, text) =>
        nl + indent + addSpan('cmt', '#' + m.slice(nl.length + indent.length + 1)));
    }
    if (['html','xml'].includes(normalized)) {
      result = result.replace(/&lt;!--[\s\S]*?--&gt;/g, m => addSpan('cmt', m));
    }
    if (normalized === 'sql') {
      result = result.replace(/--[^\n]*/g, m => addSpan('cmt', m));
    }

    // Numbers
    result = result.replace(/\b(0x[\da-fA-F]+|\d+\.?\d*(?:[eE][+-]?\d+)?[fFdDlL]?)\b/g, m => addSpan('num', m));

    // Keywords
    const kws = KEYWORDS[normalized];
    if (kws) {
      const kwRe = new RegExp(`\\b(${kws.trim().split(/\s+/).join('|')})\\b`, 'g');
      result = result.replace(kwRe, m => addSpan('kw', m));
    }

    // HTML/XML tags & attributes
    if (['html','xml'].includes(normalized)) {
      result = result.replace(/(&lt;\/?)([\w.-]+)/g, (_, lt, tag) => lt + addSpan('tag', tag));
      result = result.replace(/\s([\w:-]+)(=)/g, (_, attr, eq) => ' ' + addSpan('attr', attr) + eq);
    }

    // CSS properties
    if (normalized === 'css') {
      result = result.replace(/([\w-]+)(\s*:)/g, (_, prop, colon) => addSpan('prop', prop) + colon);
      result = result.replace(/(#[\da-fA-F]{3,8})\b/g, m => addSpan('num', m));
    }

    // JSON
    if (normalized === 'json') {
      result = result.replace(/("(?:\\.|[^"\\])*")\s*:/g, (m, key) => addSpan('key', key) + ':');
      result = result.replace(/\b(true|false|null)\b/g, m => addSpan('kw', m));
    }

    // YAML keys
    if (normalized === 'yaml') {
      result = result.replace(/^([ \t]*)([\w-]+)(\s*:)/gm, (_, ind, key, col) => ind + addSpan('key', key) + col);
    }

    // HTTP blocks
    if (normalized === 'http') {
      const METHOD_CLASS = { GET:'get', POST:'post', PUT:'put', DELETE:'delete', PATCH:'patch', HEAD:'head', OPTIONS:'options', TRACE:'options', CONNECT:'options' };
      result = result.replace(/^(GET|POST|PUT|DELETE|PATCH|HEAD|OPTIONS|TRACE|CONNECT)\b/gm,
        m => addSpan('http-' + METHOD_CLASS[m], m));
      result = result.replace(/\b(HTTP\/[\d.]+)/g, m => addSpan('kw', m));
      result = result.replace(/\b([1-5]\d{2})\b/g, m => addSpan('num', m));
      result = result.replace(/^([\w-]+)(\s*:)/gm, (_, name, colon) => addSpan('key', name) + colon);
      return result;
    }

    // Function calls
    result = result.replace(/\b([\w$]+)(\s*\()/g, (_, fn, paren) => addSpan('fn', fn) + paren);

    // Restore placeholders
    result = result.replace(/\x00P(\d+)\x00/g, (_, i) => placeholder[+i]);

    return result;
  }

  // ── Emoji Support ─────────────────────────────────────────────────────────────

  const EMOJI_MAP = {
    // Common GitHub-style emojis
    smile: '😊', laugh: '😂', cry: '😢', heart: '❤️', thumbsup: '👍', thumbsdown: '👎',
    rocket: '🚀', star: '⭐', fire: '🔥', poop: '💩', eyes: '👀', ok_hand: '👌',
    check_mark: '✅', x: '❌', warning: '⚠️', question: '❓', info: 'ℹ️', heavy_check_mark: '✔️',
    plus: '➕', minus: '➖', heavy_plus_sign: '🔼', heavy_minus_sign: '🔽',
    100: '💯', coffee: '☕', pizza: '🍕', beer: '🍺', clink: '🥂',
    cat: '🐱', dog: '🐶', mouse: '🐭', hamster: '🐹', rabbit: '🐰',
    wolf: '🐺', frog: '🐸', monkey: '🐒', gorilla: '🦍', pig: '🐷',
    pig2: '🐖', boar: '🐗', bear: '🐻', panda_face: '🐼', koala: '🐨',
    tiger: '🐯', lion: '🦁', cow: '🐮', ox: '🐂', water_buffalo: '🐃',
    chicken: '🐔', rooster: '🐓', bird: '🐦', baby_chick: '🐤', hatching_chick: '🐣',
    penguin: '🐧', snake: '🐍', turtle: '🐢', lizard: '🦎', crocodile: '🐊',
    whale: '🐳', dolphin: '🐬', fish: '🐟', octopus: '🐙', shell: '🐚',
    bug: '🐛', ant: '🐜', bee: '🐝', beetle: '🪲',
    rose: '🌹', wilted_rose: '🥀', hibiscus: '🌺', cherry_blossom: '🌸', sunflower: '🌻',
    tulip: '🌷', seedling: '🌱', evergreen_tree: '🌲', deciduous_tree: '🌳', palm_tree: '🌴',
    cactus: '🌵', ear_of_rice: '🌾', herb: '🌿', shamrock: '☘️', four_leaf_clover: '🍀',
    maple_leaf: '🍁', fallen_leaf: '🍂', leaves: '🍃',
    sun_with_face: '🌞', full_moon: '🌝', new_moon: '🌚', earth_americas: '🌎', earth_africa: '🌍',
    earth_asia: '🌏', volcano: '🌋', milky_way: '🌌', moon: '🌙', cloud: '☁️',
    snowflake: '❄️', snowman: '⛄', zap: '⚡', lightning: '🌩️', rainbow: '🌈',
    tornado: '🌪️', wave: '🌊', ocean: '🌊', droplet: '💧', sweat_drops: '💦',
    apple: '🍎', green_apple: '🍏', tangerine: '🍊', lemon: '🍋', banana: '🍌',
    watermelon: '🍉', grapes: '🍇', melon: '🍈', cherry: '🍒', peach: '🍑',
    pineapple: '🍍', kiwi: '🥝', avocado: '🥑', tomato: '🍅', eggplant: '🍆',
    carrot: '🥕', corn: '🌽', hot_pepper: '🌶️', potato: '🥔', sweet_potato: '🍠',
    bread: '🍞', cheese: '🧀', egg: '🥚', cooking: '🍳', bacon: '🥓',
    burger: '🍔', fries: '🍟', pizza: '🍕', hot_dog: '🌭', taco: '🌮',
    sandwich: '🥪', stuffed_flatbread: '🥙', salad: '🥗', shallow_pan_of_food: '🥘',
    ramen: '🍜', spaghetti: '🍝', meat_on_bone: '🍖', chicken_leg: '🍗', cut_of_meat: '🥩',
    house: '🏠', house_with_garden: '🏡', office: '🏢', post_office: '🏣', hospital: '🏥',
    bank: '🏦', hotel: '🏨', love_hotel: '🏩', convenience_store: '🏪', school: '🏫',
    department_store: '🏬', supermarket: '🏭', convenience_store: '🏪', factory: '🏭',
    japanese_castle: '🏯', european_castle: '🏰', wedding: '💒', tokyo_tower: '🗼',
    statue_of_liberty: '🗽', map: '🗺️', beach: '🏖️', desert: '🏜️', island: '🏝️',
    computer: '💻', desktop_computer: '🖥️', phone: '📱', telephone: '☎️', fax: '📠',
    camera: '📷', video_camera: '📹', television: '📺', radio: '📻', movie_camera: '🎥',
    book: '📚', books: '📖', notebook: '📓', ledger: '📒', page_facing_up: '📄',
    scroll: '📜', memo: '📝', pencil: '✏️', pen: '🖊️', crayon: '🖍️',
    briefcase: '💼', file_folder: '📁', open_file_folder: '📂', card_index: '📇',
    calendar: '📆', date: '📅', alarm_clock: '⏰', hourglass: '⌛', watch: '⌚',
    clock: '🕐', stopwatch: '⏱️', timer: '⏲️',
    musical_note: '🎵', notes: '🎶', headphone: '🎧', guitar: '🎸', saxophone: '🎷',
    trumpet: '🎺', violin: '🎻', musical_keyboard: '🎹',
    car: '🚗', taxi: '🚕', bus: '🚌', trolleybus: '🚎', race_car: '🏎️',
    police_car: '🚓', ambulance: '🚑', fire_engine: '🚒', truck: '🚚', airplane: '✈️',
    rocket: '🚀', helicopter: '🚁', steam_locomotive: '🚂', train: '🚆', tram: '🚊',
    bike: '🚲', motorcycle: '🏍️', scooter: '🛵', walking: '🚶', running: '🏃',
    dancer: '💃', man_dancing: '🕺', levitate: '🕴️', brain: '🧠', heart: '❤️',
    lungs: '🫁', anatomical_heart: '🫀', bone: '🦴', eye: '👁️', eyes: '👀',
    ear: '👂', nose: '👃', mouth: '👄', tongue: '👅', tooth: '🦷',
    handshake: '🤝', wave: '👋', thumbsup: '👍', thumbsdown: '👎', clap: '👏',
    raised_hands: '🙌', open_hands: '👐', pray: '🙏', writing_hand: '✍️', nail_care: '💅',
    lipstic: '💄', kiss: '💋', heart_eyes: '😍', sunglasses: '😎', stuck_out_tongue: '😛',
    zipper_mouth: '🤐', smirk: '😏', unamused: '😒', rolling_eyes: '🙄', flushed: '😳',
    scream: '😱', frightened: '😨', dizzy: '😵', rage: '😡', pouting: '😡',
    crying: '😢', sob: '😭', sleepy: '😪', sneezing: '🤧', mask: '😷',
    robot: '🤖', alien: '👽', ghost: '👻', skull: '💀', pile_of_poo: '💩',
    clown: '🤡', joy: '😂', rofl: '🤣', wink: '😉', innocent: '😇',
    smiling_imp: '👿', imp: '👿', japanese_goblin: '👺', japanese_ogre: '👹',
    snowman: '⛄', father_christmas: '🎅', mrs_claus: '🤶', superhero: '🦸', supervillain: '🦹',
    wizard: '🧙', witch: '🧙‍♀️', fairy: '🧚', vampire: '🧛', mermaid: '🧜‍♀️',
    elf: '🧝', genie: '🧞', zombie: '🧟', troll: '🧟',
    flag: '🏳️', flags: '🏁', checkered_flag: '🏁', triangular_flag_on_post: '🚩',
    rainbow_flag: '🏳️‍🌈', transgender_flag: '🏳️‍⚧️',
    key: '🔑', lock: '🔒', unlock: '🔓', lock_with_ink_pen: '🔏', closed_lock: '🔐',
    bell: '🔔', no_bell: '🔕', bookmark: '🔖', link: '🔗', radio_button: '🔘',
    back: '🔙', end: '🔚', on: '🔛', soon: '🔜', top: '🔝',
    gears: '⚙️', wrench: '🔧', hammer: '🔨', pick: '⛏️', hammer_and_wrench: '🛠️',
    dagger: '🗡️', crossed_swords: '⚔️', shield: '🛡️', nut_and_bolt: '🔩',
    atom: '⚛️', biohazard: '☣️', radioactive: '☢️', warning: '⚠️',
    male_sign: '♂️', female_sign: '♀️', staff_of_aesculapius: '⚕️',
    plus: '➕', minus: '➖', divide: '➗', multiplication: '✖️',
    infinity: '♾️', equals: '＝', not_equal: '≠', less_than: '＜', greater_than: '＞',
    degree: '°', ascending: '🔼', descending: '🔽',
    one: '1️⃣', two: '2️⃣', three: '3️⃣', four: '4️⃣', five: '5️⃣',
    six: '6️⃣', seven: '7️⃣', eight: '8️⃣', nine: '9️⃣', zero: '0️⃣',
    first: '1️⃣', second: '2️⃣', third: '3️⃣', fourth: '4️⃣', fifth: '5️⃣',
    sixth: '6️⃣', seventh: '7️⃣', eighth: '8️⃣', ninth: '9️⃣', tenth: '🔟',
    a: '🅰️', b: '🅱️', ab: '🆎', o2: '🅾️', sos: '🆘',
    id: '🆔', clipboard: '📋', pushpin: '📌', round_pushpin: '📍',
    pencil2: '✏️', black_nib: '✒️', heavy_check_mark: '✔️', heavy_multiplication_x: '✖️',
    x: '❌', heavy_plus_sign: '➕', heavy_minus_sign: '➖', heavy_division_sign: '➗',
    asterisk: '*️⃣', hash: '#️⃣', exclamation: '❗', question: '❓', bangbang: '‼️',
    interrobang: '⁉️', low_brightness: '🔅', high_brightness: '🔆', part_alternation_mark: '〽️'
  };

  function parseEmojis(text) {
    return text.replace(/:([a-z0-9_+-]+):/gi, (_, name) => {
      const emoji = EMOJI_MAP[name.toLowerCase()];
      return emoji ? emoji : `:${name}:`;
    });
  }

  let scrollSpyLock = false, scrollSpyTimer = null;

  function scrollToHeading(el, behavior) {
    const c = document.getElementById('md-content');
    const top = Math.min(
      Math.max(0, el.getBoundingClientRect().top - c.getBoundingClientRect().top + c.scrollTop - 60),
      c.scrollHeight - c.clientHeight
    );
    c.scrollTo({ top, behavior });
    if (behavior === 'smooth') {
      scrollSpyLock = true;
      clearTimeout(scrollSpyTimer);
      scrollSpyTimer = setTimeout(() => { scrollSpyLock = false; }, 700);
    }
  }

  // ── Math Support (KaTeX) ───────────────────────────────────────────────────────
  // katex.min.js is loaded as a content script (manifest.json) before this file,
  // so katex is available synchronously in the same isolated world.

  function renderMath(tex, displayMode) {
    try {
      return katex.renderToString(tex, {
        displayMode: displayMode,
        throwOnError: false,
        trust: true,
        strict: false
      });
    } catch (e) {
      return `<span class="math-error">${tex}</span>`;
    }
  }

  function parseMath(text) {
    // $$$...$$$ block math (process first — longest delimiter wins)
    text = text.replace(/\$\$\$([^$]+)\$\$\$/g, (_, tex) => {
      return `<div class="math-block">${renderMath(tex.trim(), true)}</div>`;
    });

    // $$...$$ block math
    text = text.replace(/\$\$([^$]+)\$\$/g, (_, tex) => {
      return `<div class="math-block">${renderMath(tex.trim(), true)}</div>`;
    });

    // $...$ inline math
    text = text.replace(/\$([^$\n]+)\$/g, (_, tex) => {
      return `<span class="math-inline">${renderMath(tex.trim(), false)}</span>`;
    });

    return text;
  }

  // ── Inline Markdown Parser ───────────────────────────────────────────────────

  function parseInline(text) {
    // Math parsing (before other processing to avoid conflicts)
    text = parseMath(text);

    // Emoji parsing
    text = parseEmojis(text);

    // Images before links
    text = text.replace(/!\[([^\]]*)\]\(([^)\s"]+)(?:\s+"([^"]*)")?\)/g, (_, alt, src, title) =>
      `<div class="img-wrap"><img src="${src}" alt="${esc(alt)}"${title ? ` title="${esc(title)}"` : ''} loading="lazy"><button class="img-download-btn" title="Download image" data-src="${esc(src)}" data-alt="${esc(alt)}"><svg width="16" height="16" viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg></button>${title ? `<div class="img-caption">${esc(title)}</div>` : ''}</div>`
    );

    // Links
    text = text.replace(/\[([^\]]+)\]\(([^)\s"]+)(?:\s+"([^"]*)")?\)/g, (_, label, href, title) => {
      const ext = /^https?:\/\//.test(href);
      return `<a href="${href}"${title ? ` title="${esc(title)}"` : ''}${ext ? ' target="_blank" rel="noopener noreferrer"' : ''}>${label}</a>`;
    });

    // Bold + Italic
    text = text.replace(/(\*\*\*|___)(.+?)\1/g, (_, __, t) => `<strong><em>${t}</em></strong>`);
    text = text.replace(/(\*\*|__)(.+?)\1/g, (_, __, t) => `<strong>${t}</strong>`);
    text = text.replace(/\*([^*\n]+?)\*/g, (_, t) => `<em>${t}</em>`);
    text = text.replace(/_([^_\n]+?)_/g, (_, t) => `<em>${t}</em>`);

    // Strikethrough, mark, sup, sub
    text = text.replace(/~~(.+?)~~/g, '<del>$1</del>');
    text = text.replace(/==(.+?)==/g, '<mark>$1</mark>');
    text = text.replace(/\^(.+?)\^/g, '<sup>$1</sup>');

    // Auto-links (bare URLs not inside href="...")
    text = text.replace(/(?<![="'>])(https?:\/\/[^\s<>")\]]+)/g,
      '<a href="$1" target="_blank" rel="noopener noreferrer">$1</a>');

    // Hard line break
    text = text.replace(/  \n|\\n/g, '<br>').replace(/\n/g, ' ');

    return text;
  }

  // ── List Parser ──────────────────────────────────────────────────────────────

  function parseList(lines, start) {
    const baseIndent = lines[start].match(/^(\s*)/)[1].length;
    const isOrdered = /^\s*\d+\./.test(lines[start]);
    const tag = isOrdered ? 'ol' : 'ul';
    const items = [];
    let i = start;

    while (i < lines.length) {
      const line = lines[i];
      if (!line.trim()) { i++; continue; }

      const indent = line.match(/^(\s*)/)[1].length;
      if (indent < baseIndent) break;

      const isUL = /^\s*[-*+]\s/.test(line);
      const isOL = /^\s*\d+\.\s/.test(line);
      if (indent === baseIndent && (isUL || isOL)) {
        let content = line.replace(/^\s*(?:[-*+]|\d+\.)\s+/, '');
        i++;

        // Continuation lines (same indent, not a new item)
        while (i < lines.length && lines[i].trim() &&
          lines[i].match(/^(\s*)/)[1].length > baseIndent &&
          !/^\s*[-*+]\s/.test(lines[i]) && !/^\s*\d+\.\s/.test(lines[i])) {
          content += ' ' + lines[i].trim();
          i++;
        }

        // Nested list
        let nested = '';
        if (i < lines.length && lines[i].trim() &&
          lines[i].match(/^(\s*)/)[1].length > baseIndent) {
          const res = parseList(lines, i);
          nested = res.html;
          i = res.next;
        }

        // Task list checkbox
        const cbMatch = content.match(/^\[([ xX])\]\s+([\s\S]+)/);
        if (cbMatch) {
          const checked = cbMatch[1].toLowerCase() === 'x' ? ' checked' : '';
          items.push(`<li class="task-item"><label><input type="checkbox" disabled${checked}> ${parseInline(cbMatch[2])}</label>${nested}</li>`);
        } else {
          items.push(`<li>${parseInline(content)}${nested}</li>`);
        }
      } else {
        break;
      }
    }

    return { html: `<${tag}>${items.join('')}</${tag}>`, next: i };
  }

  // ── Table Parser ─────────────────────────────────────────────────────────────

  function parseTable(lines) {
    if (lines.length < 2) return `<p>${parseInline(lines.join(' '))}</p>`;

    const splitRow = l => l.split('|').slice(1, -1).map(c => c.trim());
    const headers = splitRow(lines[0]);
    const aligns = splitRow(lines[1]).map(c =>
      c.startsWith(':') && c.endsWith(':') ? 'center' : c.endsWith(':') ? 'right' : 'left'
    );

    const th = headers.map((h, i) =>
      `<th style="text-align:${aligns[i] || 'left'}">${parseInline(h)}</th>`
    ).join('');

    const tbody = lines.slice(2).map(line => {
      const cells = splitRow(line);
      return '<tr>' + cells.map((c, i) =>
        `<td style="text-align:${aligns[i] || 'left'}">${parseInline(c)}</td>`
      ).join('') + '</tr>';
    }).join('');

    return `<div class="tbl-wrap"><table><thead><tr>${th}</tr></thead><tbody>${tbody}</tbody></table></div>`;
  }

  // ── Block Parser ─────────────────────────────────────────────────────────────

  function parseBlocks(src, idMap = new Map()) {
    src = src.replace(/\r\n|\r/g, '\n');
    const lines = src.split('\n');
    const out = [];
    let i = 0;

    while (i < lines.length) {
      const raw = lines[i];
      const line = raw.trimEnd();
      const trimmed = line.trim();

      // Empty line
      if (!trimmed) { i++; continue; }

      // Code block placeholder (injected by parseMarkdown)
      if (/^\x00CB\d+\x00$/.test(trimmed)) {
        out.push(trimmed);
        i++;
        continue;
      }

      // Fenced code block (shouldn't appear here, but safe guard)
      if (/^`{3,}/.test(trimmed)) { i++; continue; }

      // ATX heading
      const hm = line.match(/^(#{1,6})\s+(.*?)\s*(?:\s+#+)?\s*$/);
      if (hm) {
        const lvl = hm[1].length;
        const content = parseInline(hm[2]);
        const base = slugify(hm[2]);
        const n = idMap.get(base) || 0;
        idMap.set(base, n + 1);
        const id = n === 0 ? base : `${base}-${n}`;
        out.push(`<h${lvl} id="${id}">${content}</h${lvl}>`);
        i++;
        continue;
      }

      // Setext headings (=== or --- on next line)
      if (i + 1 < lines.length && trimmed) {
        const nxt = lines[i + 1].trim();
        if (/^=+$/.test(nxt)) {
          const base = slugify(trimmed);
          const n = idMap.get(base) || 0;
          idMap.set(base, n + 1);
          out.push(`<h1 id="${n === 0 ? base : `${base}-${n}`}">${parseInline(trimmed)}</h1>`);
          i += 2; continue;
        }
        if (/^-+$/.test(nxt) && !/^[-*+]\s/.test(trimmed)) {
          const base = slugify(trimmed);
          const n = idMap.get(base) || 0;
          idMap.set(base, n + 1);
          out.push(`<h2 id="${n === 0 ? base : `${base}-${n}`}">${parseInline(trimmed)}</h2>`);
          i += 2; continue;
        }
      }

      // Horizontal rule
      if (/^([-*_])\1{2,}\s*$/.test(trimmed)) {
        out.push('<hr>');
        i++;
        continue;
      }

      // Blockquote
      if (trimmed.startsWith('>')) {
        const bqLines = [];
        while (i < lines.length && /^\s*>/.test(lines[i])) {
          bqLines.push(lines[i].replace(/^\s*>[ \t]?/, ''));
          i++;
        }
        out.push(`<blockquote>${parseBlocks(bqLines.join('\n'), idMap)}</blockquote>`);
        continue;
      }

      // Unordered list
      if (/^\s*[-*+]\s/.test(line)) {
        const res = parseList(lines, i);
        out.push(res.html);
        i = res.next;
        continue;
      }

      // Ordered list
      if (/^\s*\d+\.\s/.test(line)) {
        const res = parseList(lines, i);
        out.push(res.html);
        i = res.next;
        continue;
      }

      // Table (current line has | and next line is a separator row)
      if (trimmed.includes('|') && i + 1 < lines.length && /^[\s|:+-]+$/.test(lines[i + 1])) {
        const tblLines = [];
        while (i < lines.length && lines[i].trim().includes('|')) {
          tblLines.push(lines[i]);
          i++;
        }
        out.push(parseTable(tblLines));
        continue;
      }

      // Paragraph — gather until blank line or block-level element
      const paraLines = [];
      while (i < lines.length) {
        const l = lines[i];
        const t = l.trim();
        if (!t) break;
        if (/^\x00CB\d+\x00$/.test(t)) break;
        if (/^#{1,6}\s/.test(l)) break;
        if (/^\s*[-*+]\s/.test(l)) break;
        if (/^\s*\d+\.\s/.test(l)) break;
        if (/^\s*>/.test(l)) break;
        if (/^([-*_])\1{2,}\s*$/.test(t)) break;
        paraLines.push(l);
        i++;
      }
      if (paraLines.length) {
        out.push(`<p>${parseInline(paraLines.join('\n'))}</p>`);
      }
    }

    return out.join('\n');
  }

  // ── Mermaid ───────────────────────────────────────────────────────────────────
  // Mermaid is loaded lazily as a <script> tag (main world) so no eval or
  // scripting permission is needed. A tiny CustomEvent bridge lets isolated-world
  // content.js send diagram code to main world and receive rendered SVG back.

  let _mermaidReady     = null; // singleton Promise — only load once per page
  let _resListenerAdded = false;
  let _diagramSeq       = 0;
  const _pending        = new Map(); // id → {resolve, reject}

  function _mermaidTheme() {
    return document.documentElement.dataset.theme === 'dark' ? 'dark' : 'default';
  }

  function ensureMermaid() {
    if (_mermaidReady) return _mermaidReady;
    _mermaidReady = new Promise((resolve, reject) => {
      // Listen for render results dispatched by the bridge (registered once).
      if (!_resListenerAdded) {
        _resListenerAdded = true;
        document.addEventListener('_mv_res', function (e) {
          const { id, svg, error } = e.detail;
          const p = _pending.get(id);
          if (!p) return;
          _pending.delete(id);
          if (error) p.reject(new Error(error));
          else p.resolve(svg);
        });
      }

      // Load both scripts as <script src> — no inline/eval, no scripting permission.
      // Resolve only after both are ready (bridge must be listening before mermaid loads).
      let bridgeOk = false, mermaidOk = false;
      const tryResolve = () => { if (bridgeOk && mermaidOk) resolve(); };

      const bridge = document.createElement('script');
      bridge.src = chrome.runtime.getURL('bridge.js');
      bridge.onload  = () => { bridgeOk = true;  tryResolve(); };
      bridge.onerror = () => { _mermaidReady = null; reject(new Error('Failed to load bridge.js')); };
      document.head.appendChild(bridge);

      const s = document.createElement('script');
      s.src = chrome.runtime.getURL('mermaid.min.js');
      s.onload  = () => { mermaidOk = true;  tryResolve(); };
      s.onerror = () => { _mermaidReady = null; reject(new Error('Failed to load mermaid.min.js')); };
      document.head.appendChild(s);
    });
    return _mermaidReady;
  }

  async function _renderOneDiagram(block) {
    const code = block.dataset.diagram;
    try {
      await ensureMermaid();
      const id = `md-diagram-${++_diagramSeq}`;
      const svg = await new Promise((resolve, reject) => {
        const timer = setTimeout(() => {
          _pending.delete(id);
          reject(new Error('Render timeout'));
        }, 15000);
        _pending.set(id, {
          resolve: v => { clearTimeout(timer); resolve(v); },
          reject:  e => { clearTimeout(timer); reject(e); },
        });
        document.dispatchEvent(new CustomEvent('_mv_req', {
          detail: { id, code, theme: _mermaidTheme() },
        }));
      });
      block.className       = 'mermaid-wrap';
      block.innerHTML       = svg;
      block.dataset.diagram = code; // keep for theme re-render
    } catch (e) {
      block.className   = 'mermaid-error';
      block.textContent = 'Diagram error: ' + e.message;
    }
  }

  function renderMermaidBlocks(container) {
    const blocks = [...container.querySelectorAll('.mermaid-pending')];
    if (!blocks.length) return;
    // Single observer for all blocks — fires as each enters the viewport.
    // 300px rootMargin gives enough lead time to finish rendering before scroll.
    const obs = new IntersectionObserver((entries, observer) => {
      entries.forEach(entry => {
        if (!entry.isIntersecting) return;
        observer.unobserve(entry.target);
        _renderOneDiagram(entry.target);
      });
    }, { rootMargin: '300px' });
    blocks.forEach(b => obs.observe(b));
  }

  // Re-render already-rendered diagrams when the user toggles dark/light mode.
  async function refreshMermaidTheme() {
    const rendered = [...document.querySelectorAll('.mermaid-wrap[data-diagram]')];
    if (!rendered.length) return;
    // Tell the bridge to re-initialize mermaid with the new theme on next render.
    document.dispatchEvent(new CustomEvent('_mv_reset_theme'));
    for (const block of rendered) await _renderOneDiagram(block);
  }

  // ── Diagram Lightbox ─────────────────────────────────────────────────────────

  (function initDiagramLightbox() {
    let lightbox = null;
    let zoom = 1;
    let svgClone = null;

    function applyZoom() {
      if (svgClone) svgClone.style.transform = `scale(${zoom})`;
    }

    function openLightbox(svgEl) {
      if (lightbox) return;
      zoom = 1;
      lightbox = document.createElement('div');
      lightbox.id = 'md-diagram-lightbox';

      svgClone = svgEl.cloneNode(true);
      svgClone.removeAttribute('width');
      svgClone.removeAttribute('height');

      const controls = document.createElement('div');
      controls.id = 'md-diagram-lightbox-controls';

      const btnOut = document.createElement('button');
      btnOut.textContent = '−';
      btnOut.title = 'Zoom out';
      btnOut.addEventListener('click', e => { e.stopPropagation(); zoom = Math.max(0.25, zoom - 0.25); applyZoom(); });

      const btnReset = document.createElement('button');
      btnReset.textContent = '1×';
      btnReset.title = 'Reset zoom';
      btnReset.addEventListener('click', e => { e.stopPropagation(); zoom = 1; applyZoom(); });

      const btnIn = document.createElement('button');
      btnIn.textContent = '+';
      btnIn.title = 'Zoom in';
      btnIn.addEventListener('click', e => { e.stopPropagation(); zoom = Math.min(4, zoom + 0.25); applyZoom(); });

      controls.append(btnOut, btnReset, btnIn);
      lightbox.append(controls, svgClone);
      document.body.appendChild(lightbox);
      document.addEventListener('keydown', onKey);
      lightbox.addEventListener('wheel', onWheel, { passive: false });
    }

    function closeLightbox() {
      if (!lightbox) return;
      lightbox.remove();
      lightbox = null;
      svgClone = null;
      document.removeEventListener('keydown', onKey);
    }

    function onKey(e) { if (e.key === 'Escape') closeLightbox(); }

    function onWheel(e) {
      e.preventDefault();
      const delta = e.deltaY > 0 ? -0.1 : 0.1;
      zoom = Math.min(4, Math.max(0.25, zoom + delta));
      applyZoom();
    }

    document.addEventListener('click', e => {
      if (lightbox) {
        if (e.target.closest('#md-diagram-lightbox-controls')) return;
        closeLightbox();
        return;
      }
      const wrap = e.target.closest('.mermaid-wrap');
      if (!wrap) return;
      const svg = wrap.querySelector('svg');
      if (svg) openLightbox(svg);
    });
  })();

  // ── Image Lightbox ─────────────────────────────────────────────────────────────

  (function initImageLightbox() {
    let lightbox = null;
    let lightboxImg = null;
    let currentIndex = 0;
    let images = [];
    let zoom = 1;

    function updateImage() {
      if (!lightboxImg || currentIndex >= images.length) return;
      lightboxImg.src = images[currentIndex].src;
      lightboxImg.alt = images[currentIndex].alt || '';

      // Update counter
      const counter = document.getElementById('md-image-lightbox-counter');
      if (counter) {
        counter.textContent = `${currentIndex + 1} / ${images.length}`;
      }

      // Update single image state
      if (images.length === 1) {
        lightbox.classList.add('single-image');
      } else {
        lightbox.classList.remove('single-image');
      }
    }

    function applyZoom() {
      if (lightboxImg) lightboxImg.style.transform = `scale(${zoom})`;
    }

    function openLightbox(imgEl, allImages, index) {
      if (lightbox) return;

      images = allImages;
      currentIndex = index;
      zoom = 1;

      lightbox = document.createElement('div');
      lightbox.id = 'md-image-lightbox';

      lightboxImg = document.createElement('img');
      lightboxImg.id = 'md-image-lightbox-img';
      lightboxImg.src = imgEl.src;
      lightboxImg.alt = imgEl.alt || '';

      const navPrev = document.createElement('button');
      navPrev.id = 'md-image-lightbox-nav-prev';
      navPrev.innerHTML = '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="15 18 9 12 15 6"/></svg>';
      navPrev.title = 'Previous image (Left arrow)';
      navPrev.setAttribute('aria-label', 'Previous image');
      navPrev.addEventListener('click', e => { e.stopPropagation(); e.preventDefault(); navigateImage(-1); });

      const navNext = document.createElement('button');
      navNext.id = 'md-image-lightbox-nav-next';
      navNext.innerHTML = '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="9 18 15 12 9 6"/></svg>';
      navNext.title = 'Next image (Right arrow)';
      navNext.setAttribute('aria-label', 'Next image');
      navNext.addEventListener('click', e => { e.stopPropagation(); e.preventDefault(); navigateImage(1); });

      const counter = document.createElement('div');
      counter.id = 'md-image-lightbox-counter';
      counter.textContent = `${currentIndex + 1} / ${images.length}`;

      const controls = document.createElement('div');
      controls.id = 'md-image-lightbox-controls';

      const btnOut = document.createElement('button');
      btnOut.textContent = '−';
      btnOut.title = 'Zoom out';
      btnOut.setAttribute('aria-label', 'Zoom out');
      btnOut.addEventListener('click', e => { e.stopPropagation(); e.preventDefault(); zoom = Math.max(0.5, zoom - 0.25); applyZoom(); });

      const btnReset = document.createElement('button');
      btnReset.textContent = '1×';
      btnReset.title = 'Reset zoom';
      btnReset.setAttribute('aria-label', 'Reset zoom');
      btnReset.addEventListener('click', e => { e.stopPropagation(); e.preventDefault(); zoom = 1; applyZoom(); });

      const btnIn = document.createElement('button');
      btnIn.textContent = '+';
      btnIn.title = 'Zoom in';
      btnIn.setAttribute('aria-label', 'Zoom in');
      btnIn.addEventListener('click', e => { e.stopPropagation(); e.preventDefault(); zoom = Math.min(3, zoom + 0.25); applyZoom(); });

      controls.append(btnOut, btnReset, btnIn);

      lightbox.append(navPrev, navNext, lightboxImg, counter, controls);
      document.body.appendChild(lightbox);

      // Trigger animation
      requestAnimationFrame(() => {
        lightbox.classList.add('visible');
      });

      document.addEventListener('keydown', onKey);
      lightbox.addEventListener('wheel', onWheel, { passive: false });
    }

    function closeLightbox() {
      if (!lightbox) return;

      lightbox.classList.remove('visible');

      setTimeout(() => {
        if (lightbox) {
          lightbox.remove();
          lightbox = null;
          lightboxImg = null;
          images = [];
          currentIndex = 0;
          zoom = 1;
          document.removeEventListener('keydown', onKey);
        }
      }, 200); // Match CSS transition
    }

    function navigateImage(direction) {
      currentIndex += direction;

      if (currentIndex < 0) {
        currentIndex = images.length - 1;
      } else if (currentIndex >= images.length) {
        currentIndex = 0;
      }

      zoom = 1; // Reset zoom when navigating
      updateImage();
      applyZoom();
    }

    function onKey(e) {
      if (e.key === 'Escape') {
        closeLightbox();
      } else if (e.key === 'ArrowLeft') {
        e.preventDefault();
        navigateImage(-1);
      } else if (e.key === 'ArrowRight') {
        e.preventDefault();
        navigateImage(1);
      }
    }

    function onWheel(e) {
      e.preventDefault();
      const delta = e.deltaY > 0 ? -0.1 : 0.1;
      zoom = Math.min(3, Math.max(0.5, zoom + delta));
      applyZoom();
    }

    // Attach click handler using document-level event delegation
    document.addEventListener('click', e => {
      const mdContent = document.getElementById('md-content');

      if (lightbox) {
        // If clicking on controls, don't close
        if (e.target.closest('#md-image-lightbox-controls') ||
            e.target.closest('#md-image-lightbox-nav-prev') ||
            e.target.closest('#md-image-lightbox-nav-next') ||
            e.target.closest('#md-image-lightbox-img')) {
          return;
        }
        closeLightbox();
        return;
      }

      // Only handle clicks within md-content
      if (!mdContent || !mdContent.contains(e.target)) return;

      const img = e.target.closest('img');
      if (!img || img.classList.contains('img-broken')) return;

      // Don't trigger lightbox when clicking download button
      if (e.target.closest('.img-download-btn')) return;

      // Collect all images in the document - be more flexible with selector
      const allImages = [...document.querySelectorAll('.md-body img:not(.img-broken)')];
      const index = allImages.indexOf(img);

      if (index !== -1) {
        e.preventDefault();
        console.log('Opening lightbox with image', index + 1, 'of', allImages.length);
        openLightbox(img, allImages, index);
      }
    });
  })();

  // ── Markdown Entry Point ─────────────────────────────────────────────────────

  function parseMarkdown(src) {
    src = src.replace(/\r\n|\r/g, '\n');

    // Store for restoring later
    const inlineCodes = [];
    const codeBlocks = [];

    // 1. Extract fenced code blocks
    src = src.replace(/^(`{3,}|~{3,})([\w+#.-]*)[^\n]*\n([\s\S]*?)\n\1\s*$/gm,
      (_, fence, lang, code) => {
        const id = codeBlocks.length;
        codeBlocks.push({ lang: lang.trim().toLowerCase(), code });
        return `\x00CB${id}\x00`;
      }
    );

    // 2. Extract inline code
    src = src.replace(/`([^`\n]+?)`/g, (_, code) => {
      const id = inlineCodes.length;
      inlineCodes.push(esc(code));
      return `\x00IC${id}\x00`;
    });

    // 3. Process blocks
    src = parseBlocks(src);

    // 4. Restore inline code
    src = src.replace(/\x00IC(\d+)\x00/g, (_, i) =>
      `<code class="ic">${inlineCodes[+i]}</code>`
    );

    // 5. Restore code blocks with highlighting
    src = src.replace(/\x00CB(\d+)\x00/g, (_, i) => {
      const { lang, code } = codeBlocks[+i];

      if (lang === 'mermaid') {
        return `<div class="mermaid-pending" data-diagram="${esc(code)}"><span class="diagram-loading">Loading diagram…</span></div>`;
      }

      const escaped = esc(code);
      const hlit = highlight(escaped, lang);
      const badge = lang ? `<span class="cb-lang">${esc(lang)}</span>` : '';
      return `<div class="cb-wrap">
  <div class="cb-bar">${badge}<button class="cb-copy">Copy</button></div>
  <pre><code class="language-${esc(lang)}">${hlit}</code></pre>
</div>`;
    });

    return src;
  }

  // ── TOC ──────────────────────────────────────────────────────────────────────

  function buildTOC(container) {
    const hs = [...container.querySelectorAll('h1,h2,h3,h4,h5,h6')];
    if (hs.length < 3) return null;

    const nav = document.createElement('nav');
    nav.id = 'md-toc';
    nav.innerHTML = '<div class="toc-title">Contents</div>';

    // Build tree from flat heading list
    const minLevel = Math.min(...hs.map(h => +h.tagName[1]));
    const root = { level: minLevel - 1, children: [] };
    const stack = [root];
    hs.forEach(h => {
      if (!h.id) return;
      const node = { level: +h.tagName[1], h, children: [] };
      while (stack.length > 1 && stack[stack.length - 1].level >= node.level) stack.pop();
      stack[stack.length - 1].children.push(node);
      stack.push(node);
    });

    const CHEVRON = `<svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="9 18 15 12 9 6"/></svg>`;

    function renderList(nodes) {
      if (!nodes.length) return null;
      const ul = document.createElement('ul');
      nodes.forEach(node => {
        const li = document.createElement('li');
        const hasChildren = node.children.length > 0;
        const row = document.createElement('div');
        row.className = 'toc-row';

        if (hasChildren) {
          li.classList.add('toc-open');
          const btn = document.createElement('button');
          btn.className = 'toc-toggle';
          btn.setAttribute('aria-label', 'Toggle section');
          btn.setAttribute('aria-expanded', 'true');
          btn.innerHTML = CHEVRON;
          btn.addEventListener('click', e => {
            e.stopPropagation();
            const open = li.classList.toggle('toc-open');
            btn.setAttribute('aria-expanded', String(open));
          });
          row.appendChild(btn);
        } else {
          const leaf = document.createElement('span');
          leaf.className = 'toc-leaf';
          row.appendChild(leaf);
        }

        const a = document.createElement('a');
        a.href = '#' + node.h.id;
        a.textContent = node.h.textContent;
        row.appendChild(a);
        li.appendChild(row);

        if (hasChildren) {
          const childUl = renderList(node.children);
          if (childUl) li.appendChild(childUl);
        }
        ul.appendChild(li);
      });
      return ul;
    }

    const list = renderList(root.children);
    if (list) nav.appendChild(list);
    return nav;
  }

  // ── State ────────────────────────────────────────────────────────────────────

  let rawContent = '';
  let currentMode = 'view'; // 'view' | 'edit' | 'split'
  let previewTimer = null;

  // ── Helpers ──────────────────────────────────────────────────────────────────

  function getDocTitle(src) {
    const m = src.match(/^#\s+(.+)/m);
    return m ? m[1].replace(/[*_`]/g, '') : getFilename();
  }

  function applyTheme(theme) {
    const root = document.documentElement;
    if (theme === 'auto') root.removeAttribute('data-theme');
    else root.dataset.theme = theme;
  }

  function attachCopyButtons(container) {
    container.querySelectorAll('.cb-copy').forEach(btn => {
      btn.addEventListener('click', () => {
        const code = btn.closest('.cb-wrap').querySelector('code').textContent;
        navigator.clipboard.writeText(code).then(() => {
          btn.textContent = 'Copied!';
          btn.classList.add('copied');
          setTimeout(() => { btn.textContent = 'Copy'; btn.classList.remove('copied'); }, 2000);
        }).catch(() => {});
      });
    });
  }

  function attachImageDownloadButtons(container) {
    container.querySelectorAll('.img-download-btn').forEach(btn => {
      btn.addEventListener('click', async (e) => {
        e.preventDefault();
        e.stopPropagation();

        const src = btn.dataset.src;
        const alt = btn.dataset.alt || 'image';
        const img = btn.closest('.img-wrap').querySelector('img');

        if (!src) return;

        try {
          // Try to fetch the image
          const response = await fetch(src);
          const blob = await response.blob();

          // Create download link
          const url = URL.createObjectURL(blob);
          const link = document.createElement('a');
          link.href = url;

          // Extract filename from URL or use alt text
          let filename = alt;
          try {
            const urlObj = new URL(src, location.href);
            const pathname = urlObj.pathname;
            const match = pathname.match(/\/([^\/]+\.(png|jpg|jpeg|gif|webp|svg|bmp|ico))$/i);
            if (match) filename = match[1];
          } catch {}

          // Ensure filename has extension
          if (!filename.match(/\.(png|jpg|jpeg|gif|webp|svg|bmp|ico)$/i)) {
            filename += '.png';
          }

          link.download = filename;
          document.body.appendChild(link);
          link.click();
          document.body.removeChild(link);

          // Clean up object URL
          setTimeout(() => URL.revokeObjectURL(url), 100);

          // Show feedback
          const originalHTML = btn.innerHTML;
          btn.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24"><polyline points="20 6 9 17 4 12"/></svg>';
          btn.classList.add('downloaded');
          setTimeout(() => {
            btn.innerHTML = originalHTML;
            btn.classList.remove('downloaded');
          }, 2000);

        } catch (error) {
          console.error('Failed to download image:', error);
          // Fallback: open in new tab
          window.open(src, '_blank');
        }
      });
    });
  }

  function showToast(msg) {
    const toast = document.getElementById('md-toast');
    if (!toast) return;
    toast.textContent = msg;
    toast.classList.add('visible');
    clearTimeout(toast._t);
    toast._t = setTimeout(() => toast.classList.remove('visible'), 2200);
  }

  // ── Live Preview ─────────────────────────────────────────────────────────────

  function schedulePreviewUpdate() {
    clearTimeout(previewTimer);
    previewTimer = setTimeout(flushPreview, 220);
  }

  function attachImageErrorHandlers(container) {
    container.querySelectorAll('img').forEach(img => {
      if (img.complete && img.naturalWidth === 0 && img.src) markImageBroken(img);
      else img.addEventListener('error', () => markImageBroken(img), { once: true });
    });
  }

  function markImageBroken(img) {
    const span = document.createElement('span');
    span.className = 'img-broken';
    span.textContent = '⚠ Image not found: ' + (img.alt || img.getAttribute('src') || 'unknown');
    img.replaceWith(span);
  }

  function flushPreview() {
    const ta = document.getElementById('md-editor');
    const body = document.querySelector('#md-content .md-body');
    if (!ta || !body) return;
    rawContent = ta.value;
    body.innerHTML = parseMarkdown(rawContent);
    attachCopyButtons(body);
    renderMermaidBlocks(body);
    attachImageErrorHandlers(body);
    attachImageDownloadButtons(body);
  }

  // ── Search Functionality ───────────────────────────────────────────────────────

  let searchState = {
    query: '',
    matches: [],
    currentMatch: -1,
    highlights: []
  };

  function clearHighlights() {
    searchState.highlights.forEach(el => {
      const parent = el.parentNode;
      if (parent) {
        parent.replaceChild(document.createTextNode(el.textContent), el);
        parent.normalize();
      }
    });
    searchState.highlights = [];
    searchState.matches = [];
    searchState.currentMatch = -1;
  }

  function highlightSearchMatches(query, container) {
    clearHighlights();

    if (!query.trim()) {
      updateSearchUI();
      return;
    }

    const regex = new RegExp(`(${escapeRegex(query)})`, 'gi');
    const textNodes = [];

    // Find all text nodes in the content
    const walker = document.createTreeWalker(
      container,
      NodeFilter.SHOW_TEXT,
      {
        acceptNode: (node) => {
          if (node.parentElement.closest('code, pre, .no-search, .cb-wrap, .mermaid-wrap')) {
            return NodeFilter.FILTER_REJECT;
          }
          return node.textContent.trim() ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP;
        }
      }
    );

    let node;
    while (node = walker.nextNode()) {
      if (regex.test(node.textContent)) {
        textNodes.push(node);
      }
    }

    // Apply highlighting
    textNodes.forEach(node => {
      const parent = node.parentNode;
      const fragment = document.createDocumentFragment();
      const parts = node.textContent.split(regex);

      parts.forEach(part => {
        if (regex.test(part)) {
          const mark = document.createElement('mark');
          mark.className = 'search-highlight';
          mark.textContent = part;
          searchState.matches.push(mark);
          fragment.appendChild(mark);
          searchState.highlights.push(mark);
        } else {
          fragment.appendChild(document.createTextNode(part));
        }
      });

      parent.replaceChild(fragment, node);
    });

    searchState.query = query;
    if (searchState.matches.length > 0) {
      searchState.currentMatch = 0;
      scrollToMatch(0);
    }
    updateSearchUI();
  }

  function escapeRegex(string) {
    return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  function scrollToMatch(index) {
    if (index < 0 || index >= searchState.matches.length) return;

    // Remove previous current match highlight
    searchState.matches.forEach(m => m.classList.remove('search-highlight-current'));

    // Add current match highlight
    const current = searchState.matches[index];
    current.classList.add('search-highlight-current');

    // Scroll into view
    current.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }

  function updateSearchUI() {
    const count = document.getElementById('search-count');
    const input = document.getElementById('search-input');

    if (count) {
      if (searchState.matches.length === 0) {
        count.textContent = 'No results';
        count.className = 'search-count search-no-results';
      } else {
        count.textContent = `${searchState.currentMatch + 1} of ${searchState.matches.length}`;
        count.className = 'search-count';
      }
    }
  }

  function nextMatch() {
    if (searchState.matches.length === 0) return;
    searchState.currentMatch = (searchState.currentMatch + 1) % searchState.matches.length;
    scrollToMatch(searchState.currentMatch);
    updateSearchUI();
  }

  function prevMatch() {
    if (searchState.matches.length === 0) return;
    searchState.currentMatch = (searchState.currentMatch - 1 + searchState.matches.length) % searchState.matches.length;
    scrollToMatch(searchState.currentMatch);
    updateSearchUI();
  }

  function closeSearch() {
    const searchContainer = document.getElementById('search-container');
    if (searchContainer) {
      searchContainer.classList.add('search-hidden');
    }
    clearHighlights();
  }

  function openSearch() {
    const searchContainer = document.getElementById('search-container');
    const searchInput = document.getElementById('search-input');
    if (searchContainer) {
      searchContainer.classList.remove('search-hidden');
      searchInput.focus();
    }
  }

  // ── Word Count ───────────────────────────────────────────────────────────────

  function updateWordCount() {
    const ta = document.getElementById('md-editor');
    const wc = document.getElementById('md-wc');
    if (!ta || !wc) return;
    const plain = ta.value.replace(/```[\s\S]*?```/g, '').replace(/`[^`]+`/g, '');
    const words = plain.trim() ? plain.trim().split(/\s+/).length : 0;
    wc.textContent = `${words.toLocaleString()} words · ${ta.value.length.toLocaleString()} chars`;
  }

  // ── Save / Download ──────────────────────────────────────────────────────────

  async function saveFile() {
    const ta = document.getElementById('md-editor');
    const content = ta ? ta.value : rawContent;
    if (typeof window.showSaveFilePicker === 'function') {
      try {
        const handle = await window.showSaveFilePicker({
          suggestedName: getFilename(),
          types: [{ description: 'Markdown', accept: { 'text/markdown': ['.md', '.markdown'] } }]
        });
        const writable = await handle.createWritable();
        await writable.write(content);
        await writable.close();
        showToast('File saved');
        return;
      } catch (e) {
        if (e.name === 'AbortError') return;
      }
    }
    downloadFile(content);
  }

  function downloadFile(content) {
    const blob = new Blob([content], { type: 'text/markdown;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    Object.assign(document.createElement('a'), { href: url, download: getFilename() }).click();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
    showToast('Downloaded');
  }

  // ── Export to DOCX ────────────────────────────────────────────────────────────

  const EXPORT_MD_EXT = /\.(md|markdown|mdown|mkd|mkdn|mdwn|mdtext)$/i;

  async function doExport() {
    const btn = document.getElementById('btn-export');
    if (!btn || btn.classList.contains('exporting')) return;
    if (!window.MdDocx) { showToast('Export unavailable'); return; }
    if (currentMode === 'edit' || currentMode === 'split') flushPreview();

    const container = document.querySelector('#md-content .md-body');
    if (!container) { showToast('Nothing to export'); return; }

    const orig = btn.innerHTML;
    btn.classList.add('exporting');
    btn.disabled = true;
    btn.innerHTML = '<svg class="spin" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 12a9 9 0 1 1-6.219-8.56"/></svg><span class="btn-label">Exporting…</span>';

    // Diagrams bake in the current theme; force light so exported PNGs match the light doc.
    const root = document.documentElement;
    const prevTheme = root.dataset.theme;
    const forceLight = prevTheme === 'dark' && !!document.querySelector('.mermaid-wrap[data-diagram]');
    if (forceLight) { root.dataset.theme = 'light'; await refreshMermaidTheme(); }

    try {
      const base = (getFilename().replace(EXPORT_MD_EXT, '') || 'document');
      const { skipped } = await window.MdDocx.export({
        container,
        filename: base + '.docx',
        title: document.title || base,
      });
      showToast(skipped
        ? `Exported — ${skipped} item${skipped > 1 ? 's' : ''} couldn't be embedded`
        : 'Exported ' + base + '.docx');
    } catch (err) {
      console.error('DOCX export failed:', err);
      showToast('Export failed — see console');
    } finally {
      if (forceLight) { root.dataset.theme = prevTheme; await refreshMermaidTheme(); }
      btn.classList.remove('exporting');
      btn.disabled = false;
      btn.innerHTML = orig;
    }
  }

  // ── Format Apply ─────────────────────────────────────────────────────────────

  function applyFormat(type) {
    const ta = document.getElementById('md-editor');
    if (!ta) return;

    const ss = ta.selectionStart;
    const se = ta.selectionEnd;
    const val = ta.value;
    const sel = val.slice(ss, se);
    const lineStart = val.lastIndexOf('\n', ss - 1) + 1;
    const lineEndRaw = val.indexOf('\n', se);
    const lineEnd = lineEndRaw === -1 ? val.length : lineEndRaw;
    const currentLine = val.slice(lineStart, lineEnd);

    let newVal, newSS, newSE;

    const wrap = (before, placeholder, after = before) => {
      const text = sel || placeholder;
      newVal = val.slice(0, ss) + before + text + after + val.slice(se);
      newSS = ss + before.length;
      newSE = newSS + text.length;
    };

    const wrapLines = (prefix, placeholder) => {
      const lines = (sel || placeholder).split('\n');
      const ins = lines.map((l, i) =>
        typeof prefix === 'function' ? prefix(l, i) : prefix + l
      ).join('\n');
      newVal = val.slice(0, ss) + ins + val.slice(se);
      newSS = ss; newSE = ss + ins.length;
    };

    switch (type) {
      case 'bold':   wrap('**', 'bold text');   break;
      case 'italic': wrap('*',  'italic text'); break;
      case 'strike': wrap('~~', 'text');        break;
      case 'code':   wrap('`',  'code');        break;

      case 'link': {
        const label = sel || 'link text';
        const ins = `[${label}](url)`;
        newVal = val.slice(0, ss) + ins + val.slice(se);
        newSS = ss + label.length + 3; newSE = newSS + 3; // select "url"
        break;
      }
      case 'image': {
        const alt = sel || 'alt text';
        const ins = `![${alt}](url)`;
        newVal = val.slice(0, ss) + ins + val.slice(se);
        newSS = ss + alt.length + 4; newSE = newSS + 3;
        break;
      }

      case 'h1': case 'h2': case 'h3': case 'h4': {
        const lvl = { h1: '#', h2: '##', h3: '###', h4: '####' }[type];
        const stripped = currentLine.replace(/^#{1,6}\s/, '');
        const existing = currentLine.match(/^(#{1,6})\s/);
        const newLine = (existing && existing[1] === lvl) ? stripped : `${lvl} ${stripped}`;
        newVal = val.slice(0, lineStart) + newLine + val.slice(lineEnd);
        newSS = lineStart + newLine.length; newSE = newSS;
        break;
      }

      case 'ul':    wrapLines('- ',    'List item');  break;
      case 'ol':    wrapLines((l, i) => `${i + 1}. ${l}`, 'List item'); break;
      case 'task':  wrapLines('- [ ] ', 'Task item'); break;
      case 'quote': wrapLines('> ',    'Quoted text'); break;

      case 'codeblock': {
        const code = sel || 'code here';
        const ins = `\`\`\`\n${code}\n\`\`\``;
        newVal = val.slice(0, ss) + ins + val.slice(se);
        newSS = ss + 4; newSE = newSS + code.length;
        break;
      }
      case 'hr': {
        const prefix = ss > 0 && val[ss - 1] !== '\n' ? '\n' : '';
        const ins = `${prefix}\n---\n\n`;
        newVal = val.slice(0, ss) + ins + val.slice(se);
        newSS = newSE = ss + ins.length;
        break;
      }
      case 'table': {
        const ins = '\n| Header 1 | Header 2 | Header 3 |\n| :--- | :--- | :--- |\n| Cell | Cell | Cell |\n';
        newVal = val.slice(0, ss) + ins + val.slice(se);
        newSS = newSE = ss + ins.length;
        break;
      }
      default: return;
    }

    ta.value = newVal;
    ta.selectionStart = newSS;
    ta.selectionEnd = newSE;
    ta.focus();
    schedulePreviewUpdate();
    updateWordCount();
  }

  // ── Editor Keyboard Handler ───────────────────────────────────────────────────

  function handleEditorKeydown(e) {
    const ta = e.currentTarget;

    // Ctrl / Cmd shortcuts
    if (e.ctrlKey || e.metaKey) {
      const map = { b: 'bold', i: 'italic', k: 'link', e: 'code' };
      if (map[e.key.toLowerCase()]) { e.preventDefault(); applyFormat(map[e.key.toLowerCase()]); return; }
      if (e.key.toLowerCase() === 's') { e.preventDefault(); saveFile(); return; }
    }

    // Tab → indent / unindent
    if (e.key === 'Tab') {
      e.preventDefault();
      const ss = ta.selectionStart;
      const se = ta.selectionEnd;
      const val = ta.value;
      if (e.shiftKey) {
        const ls = val.lastIndexOf('\n', ss - 1) + 1;
        if (val.slice(ls, ls + 2) === '  ') {
          ta.value = val.slice(0, ls) + val.slice(ls + 2);
          ta.selectionStart = Math.max(ls, ss - 2);
          ta.selectionEnd   = Math.max(ls, se - 2);
        }
      } else {
        ta.value = val.slice(0, ss) + '  ' + val.slice(se);
        ta.selectionStart = ta.selectionEnd = ss + 2;
      }
      schedulePreviewUpdate();
      return;
    }

    // Enter → continue list items automatically
    if (e.key === 'Enter') {
      const ss = ta.selectionStart;
      const val = ta.value;
      const ls = val.lastIndexOf('\n', ss - 1) + 1;
      const line = val.slice(ls, ss);

      const ul = line.match(/^(\s*)([-*+])\s(?:\[[ xX]\]\s)?(.*)$/);
      if (ul) {
        e.preventDefault();
        if (!ul[3].trim()) {
          // Empty item → exit list
          ta.value = val.slice(0, ls) + '\n' + val.slice(ss);
          ta.selectionStart = ta.selectionEnd = ls + 1;
        } else {
          const prefix = `${ul[1]}${ul[2]} `;
          ta.value = val.slice(0, ss) + '\n' + prefix + val.slice(ss);
          ta.selectionStart = ta.selectionEnd = ss + 1 + prefix.length;
        }
        schedulePreviewUpdate(); return;
      }

      const ol = line.match(/^(\s*)(\d+)\.\s(.*)$/);
      if (ol) {
        e.preventDefault();
        if (!ol[3].trim()) {
          ta.value = val.slice(0, ls) + '\n' + val.slice(ss);
          ta.selectionStart = ta.selectionEnd = ls + 1;
        } else {
          const prefix = `${ol[1]}${+ol[2] + 1}. `;
          ta.value = val.slice(0, ss) + '\n' + prefix + val.slice(ss);
          ta.selectionStart = ta.selectionEnd = ss + 1 + prefix.length;
        }
        schedulePreviewUpdate(); return;
      }
    }
  }

  // ── Scroll Sync (split mode) ──────────────────────────────────────────────────

  function setupScrollSync() {
    const editor  = document.getElementById('md-editor');
    const preview = document.getElementById('md-content');
    if (!editor || !preview) return;
    let lock = false;
    editor.addEventListener('scroll', () => {
      if (lock) return; lock = true;
      const r = editor.scrollTop / Math.max(1, editor.scrollHeight - editor.clientHeight);
      preview.scrollTop = r * (preview.scrollHeight - preview.clientHeight);
      requestAnimationFrame(() => { lock = false; });
    });
    preview.addEventListener('scroll', () => {
      if (lock) return; lock = true;
      const r = preview.scrollTop / Math.max(1, preview.scrollHeight - preview.clientHeight);
      editor.scrollTop = r * (editor.scrollHeight - editor.clientHeight);
      requestAnimationFrame(() => { lock = false; });
    });
  }

  // ── Mode Sync Helpers ─────────────────────────────────────────────────────────

  // Returns a stateful function that generates deduplicated heading IDs —
  // mirrors the counter logic in parseBlocks so editor-text scanning stays in sync.
  function makeIdCounter() {
    const map = new Map();
    return text => {
      const base = slugify(text);
      const n = map.get(base) || 0;
      map.set(base, n + 1);
      return n === 0 ? base : `${base}-${n}`;
    };
  }

  // Returns the ID of the heading currently visible near the top of the preview.
  function getActiveHeadingId() {
    const active = document.querySelector('#md-toc a.active');
    if (active) return active.getAttribute('href').slice(1);
    const content = document.getElementById('md-content');
    if (!content) return null;
    const top = content.getBoundingClientRect().top;
    for (const h of content.querySelectorAll('h1,h2,h3,h4,h5,h6')) {
      const r = h.getBoundingClientRect();
      if (r.top >= top && r.top < top + content.clientHeight * 0.6) return h.id;
    }
    return null;
  }

  // Scrolls the editor textarea so that the heading with `id` is near the top.
  function syncEditorToHeading(id) {
    const editor = document.getElementById('md-editor');
    if (!editor || !id) return;
    const lines = editor.value.split('\n');
    const nextId = makeIdCounter();
    for (let i = 0; i < lines.length; i++) {
      const atx = lines[i].match(/^(#{1,6})\s+(.*?)\s*(?:\s+#+)?\s*$/);
      if (atx) {
        if (nextId(atx[2]) === id) {
          const lh = parseFloat(getComputedStyle(editor).lineHeight) || 22;
          editor.scrollTop = Math.max(0, i * lh - editor.clientHeight * 0.1);
          return;
        }
        continue;
      }
      if (i + 1 < lines.length && lines[i].trim()) {
        const nxt = lines[i + 1].trim();
        if (/^=+$/.test(nxt) || (/^-+$/.test(nxt) && !/^[-*+]\s/.test(lines[i]))) {
          if (nextId(lines[i].trim()) === id) {
            const lh = parseFloat(getComputedStyle(editor).lineHeight) || 22;
            editor.scrollTop = Math.max(0, i * lh - editor.clientHeight * 0.1);
            return;
          }
        }
      }
    }
  }


  // ── Mode Management ───────────────────────────────────────────────────────────

  function setMode(mode) {
    const prevMode = currentMode;
    currentMode = mode;
    const layout   = document.getElementById('md-layout');
    const tocWrap  = document.getElementById('md-toc-wrap');
    const editor   = document.getElementById('md-editor');

    if (!layout) return;

    // ── Capture scroll anchors BEFORE layout change (elements still visible) ──
    // View → Edit/Split: remember which heading is on screen in the preview
    const viewHeadingId = (mode !== 'view' && prevMode === 'view')
      ? getActiveHeadingId() : null;

    // Edit/Split → View: read editor position while editor is still displayed
    let editHeadingId = null;
    let editScrollRatio = null;
    if (mode === 'view' && (prevMode === 'edit' || prevMode === 'split')) {
      if (editor) {
        const lh = parseFloat(getComputedStyle(editor).lineHeight) || 22;
        const topLine = Math.floor(editor.scrollTop / lh);
        const lines = editor.value.split('\n');
        const nextId = makeIdCounter();
        // Scan forward so the dedup counter matches parseBlocks order.
        // Record every heading at or above topLine+1; last one found is nearest.
        for (let i = 0; i <= Math.min(topLine + 1, lines.length - 1); i++) {
          const atx = lines[i].match(/^(#{1,6})\s+(.*?)\s*(?:\s+#+)?\s*$/);
          if (atx) { editHeadingId = nextId(atx[2]); continue; }
          if (i + 1 < lines.length && lines[i].trim()) {
            const nxt = lines[i + 1].trim();
            if (/^=+$/.test(nxt) || (/^-+$/.test(nxt) && !/^[-*+]\s/.test(lines[i]))) {
              editHeadingId = nextId(lines[i].trim());
            }
          }
        }
        editScrollRatio = editor.scrollTop / Math.max(1, editor.scrollHeight - editor.clientHeight);
      }
    }

    // Update mode button states
    document.querySelectorAll('.mode-btn').forEach(b =>
      b.classList.toggle('active', b.dataset.mode === mode)
    );

    layout.dataset.mode = mode;

    if (mode === 'view') {
      // Apply captured editor position to the now-visible preview
      if (prevMode === 'edit' || prevMode === 'split') {
        requestAnimationFrame(() => {
          const content = document.getElementById('md-content');
          if (!content) return;
          if (editHeadingId) {
            const el = content.querySelector('#' + CSS.escape(editHeadingId));
            if (el) { scrollToHeading(el, 'instant'); return; }
          }
          content.scrollTop = editScrollRatio * (content.scrollHeight - content.clientHeight);
        });
      }
    }

    if (mode === 'edit' || mode === 'split') {
      // Seed editor with current rawContent on first use
      if (editor && !editor.dataset.seeded) {
        editor.value = rawContent;
        editor.dataset.seeded = '1';
        updateWordCount();
      }
      if (mode === 'split') {
        flushPreview();
        setupScrollSync();
      }
      editor && editor.focus();
      // Scroll editor to the section that was visible in the preview
      if (viewHeadingId) requestAnimationFrame(() => syncEditorToHeading(viewHeadingId));
    }
  }

  // ── Format Bar HTML ───────────────────────────────────────────────────────────

  function buildFmtBarHTML() {
    const btn = (fmt, label, title) =>
      `<button class="fmt-btn" data-fmt="${fmt}" title="${title}">${label}</button>`;
    const sep = '<span class="fmt-sep"></span>';

    return `
      <div class="fmt-group">
        ${btn('h1','H1','Heading 1')}${btn('h2','H2','Heading 2')}${btn('h3','H3','Heading 3')}
      </div>${sep}
      <div class="fmt-group">
        ${btn('bold',  '<b>B</b>',    'Bold (Ctrl+B)')}
        ${btn('italic','<i>I</i>',    'Italic (Ctrl+I)')}
        ${btn('strike','<s>S</s>',    'Strikethrough')}
      </div>${sep}
      <div class="fmt-group">
        ${btn('link',  '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/></svg> Link', 'Link (Ctrl+K)')}
        ${btn('image', '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><polyline points="21 15 16 10 5 21"/></svg> Image', 'Image')}
      </div>${sep}
      <div class="fmt-group">
        ${btn('code',      '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="16 18 22 12 16 6"/><polyline points="8 6 2 12 8 18"/></svg>', 'Inline code (Ctrl+E)')}
        ${btn('codeblock', '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><rect x="2" y="3" width="20" height="18" rx="2"/><line x1="2" y1="9" x2="22" y2="9"/><line x1="8" y1="15" x2="10" y2="15"/><line x1="14" y1="15" x2="16" y2="15"/></svg>', 'Code block')}
      </div>${sep}
      <div class="fmt-group">
        ${btn('ul',   '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><line x1="9" y1="6" x2="20" y2="6"/><line x1="9" y1="12" x2="20" y2="12"/><line x1="9" y1="18" x2="20" y2="18"/><circle cx="4" cy="6" r="1.5" fill="currentColor"/><circle cx="4" cy="12" r="1.5" fill="currentColor"/><circle cx="4" cy="18" r="1.5" fill="currentColor"/></svg>', 'Bullet list')}
        ${btn('ol',   '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><line x1="10" y1="6" x2="21" y2="6"/><line x1="10" y1="12" x2="21" y2="12"/><line x1="10" y1="18" x2="21" y2="18"/><path d="M4 6h1v4"/><path d="M4 10h2"/><path d="M6 18H4c0-1 2-2 2-3s-1-1.5-2-1"/></svg>', 'Numbered list')}
        ${btn('task', '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><rect x="3" y="5" width="6" height="6" rx="1"/><polyline points="5 8 7 10 9 6"/><line x1="13" y1="8" x2="21" y2="8"/><rect x="3" y="13" width="6" height="6" rx="1"/><line x1="13" y1="16" x2="21" y2="16"/></svg>', 'Task list')}
      </div>${sep}
      <div class="fmt-group">
        ${btn('quote', '<svg width="13" height="13" viewBox="0 0 24 24" fill="currentColor"><path d="M3 21c3 0 7-1 7-8V5c0-1.25-.756-2.017-2-2H4c-1.25 0-2 .75-2 1.972V11c0 1.25.75 2 2 2 1 0 1 0 1 1v1c0 1-1 2-2 2s-1 .008-1 1.031V20c0 1 0 1 1 1zm12 0c3 0 7-1 7-8V5c0-1.25-.757-2.017-2-2h-4c-1.25 0-2 .75-2 1.972V11c0 1.25.75 2 2 2h.75c0 2.25.25 4-2.75 4v3c0 1 0 1 1 1z"/></svg>', 'Blockquote')}
        ${btn('hr',    '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><line x1="5" y1="12" x2="19" y2="12"/></svg>', 'Horizontal rule')}
        ${btn('table', '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="3" y1="15" x2="21" y2="15"/><line x1="12" y1="3" x2="12" y2="21"/></svg>', 'Insert table')}
      </div>
      <span class="fmt-spacer"></span>
      <span id="md-wc" class="wc-label"></span>
      <button id="btn-save" class="fmt-action" title="Save file (Ctrl+S)">
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg> Save
      </button>
      <button id="btn-dl" class="fmt-action" title="Download as .md">
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg> Download
      </button>`;
  }

  // ── Page Renderer ─────────────────────────────────────────────────────────────

  function render(src) {
    rawContent = src;
    const title = getDocTitle(src);
    const html  = parseMarkdown(src);

    document.title = title;
    document.head.innerHTML = `
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>${esc(title)}</title>
      <link rel="stylesheet" href="${chrome.runtime.getURL('katex/katex.min.css')}">`;

    document.body.innerHTML = `
      <div id="md-root">
        <header id="md-bar">
          <div class="bar-left">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/></svg>
            <span id="md-fname">${esc(getFilename())}</span>
          </div>
          <div class="bar-center">
            <div class="mode-sw">
              <button class="mode-btn active" data-mode="view">
                <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg> View
              </button>
              <button class="mode-btn" data-mode="edit">
                <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg> Edit
              </button>
              <button class="mode-btn" data-mode="split">
                <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="12" y1="3" x2="12" y2="21"/></svg> Split
              </button>
            </div>
          </div>
          <div class="bar-right">
            <span class="nav-badge" title="Mermaid diagram support enabled">
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 2L2 7l10 5 10-5-10-5z"/><path d="M2 17l10 5 10-5"/><path d="M2 12l10 5 10-5"/></svg>
              Mermaid
            </span>
            <button
              id="btn-export"
              title="Export to Word (.docx) — Ctrl+Shift+E"
              aria-label="Export to DOCX"
            >
              <svg
                width="15"
                height="15"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                stroke-width="2"
              >
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
                <polyline points="14 2 14 8 20 8" />
                <path d="M8 13l1.5 4 1.5-3 1.5 3 1.5-4" />
              </svg>
              <span class="btn-label">DOCX</span>
            </button>
            <button id="btn-search" title="Search (Ctrl+F)" aria-label="Search">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
            </button>
            <button id="btn-theme" title="Toggle dark / light mode" aria-label="Toggle theme">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/></svg>
            </button>
          </div>
        </header>

        <div id="search-container" class="search-hidden">
          <div class="search-bar">
            <div class="search-input-wrapper">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
              <input type="text" id="search-input" placeholder="Search in document..." autocomplete="off">
              <button id="search-close" title="Close" aria-label="Close search">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
              </button>
            </div>
            <div class="search-nav">
              <button id="search-prev" title="Previous match (Shift+Enter)" aria-label="Previous match">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="15 18 9 12 15 6"/></svg>
              </button>
              <span id="search-count" class="search-count">No results</span>
              <button id="search-next" title="Next match (Enter)" aria-label="Next match">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="9 18 15 12 9 6"/></svg>
              </button>
            </div>
          </div>
        </div>

        <div id="md-layout" data-mode="view">
          <aside id="md-toc-wrap" hidden>
            <button id="btn-toc" title="Collapse sidebar" aria-label="Toggle sidebar">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="9" y1="3" x2="9" y2="21"/></svg>
            </button>
          </aside>
          <main id="md-content"><div class="md-body">${html}</div></main>

          <div id="md-editor-pane">
            <div id="md-fmt-bar">${buildFmtBarHTML()}</div>
            <textarea id="md-editor" spellcheck="true" placeholder="Start writing Markdown…"></textarea>
          </div>

        </div>
      </div>
      <div id="md-toast"></div>`;

    // ── TOC ──────────────────────────────────────────────────────────────────
    const tocWrap = document.getElementById('md-toc-wrap');
    const toc = buildTOC(document.getElementById('md-content'));
    if (toc) { tocWrap.appendChild(toc); tocWrap.removeAttribute('hidden'); }
    setupScrollSpy();

    // ── Restore scroll position after reload ──────────────────────────────────
    const scrollKey = 'md-scroll:' + location.pathname + location.search;
    const content = document.getElementById('md-content');
    const saved = sessionStorage.getItem(scrollKey);
    if (saved) content.scrollTop = parseInt(saved, 10);
    content.addEventListener('scroll', () => {
      sessionStorage.setItem(scrollKey, content.scrollTop);
    }, { passive: true });

    // ── Copy buttons, Mermaid diagrams & image errors (initial render) ───────
    const mdContent = document.getElementById('md-content');
    attachCopyButtons(mdContent);
    renderMermaidBlocks(mdContent);
    attachImageErrorHandlers(mdContent);
    attachImageDownloadButtons(mdContent);

    // ── Mode switcher ────────────────────────────────────────────────────────
    document.querySelectorAll('.mode-btn').forEach(btn => {
      btn.addEventListener('click', () => setMode(btn.dataset.mode));
    });

    // ── Format toolbar ────────────────────────────────────────────────────────
    document.getElementById('md-fmt-bar').addEventListener('click', e => {
      const btn = e.target.closest('.fmt-btn');
      if (btn) applyFormat(btn.dataset.fmt);
    });
    document.getElementById('btn-save').addEventListener('click', saveFile);
    document.getElementById('btn-dl').addEventListener('click', () => {
      const ta = document.getElementById('md-editor');
      downloadFile(ta ? ta.value : rawContent);
    });

    // ── Export to DOCX ────────────────────────────────────────────────────────
    document.getElementById('btn-export').addEventListener('click', doExport);
    document.addEventListener('keydown', (e) => {
      if ((e.ctrlKey || e.metaKey) && e.shiftKey && (e.key === 'E' || e.key === 'e')) {
        e.preventDefault();
        doExport();
      }
    });

    // ── Editor textarea ───────────────────────────────────────────────────────
    const editorEl = document.getElementById('md-editor');
    editorEl.addEventListener('keydown', handleEditorKeydown);
    editorEl.addEventListener('input', () => {
      schedulePreviewUpdate();
      updateWordCount();
    });

    // ── Sidebar toggle (btn lives inside the sidebar, always visible) ────────
    document.getElementById('btn-toc').addEventListener('click', () => {
      const collapsed = tocWrap.classList.toggle('toc-collapsed');
      const btn = document.getElementById('btn-toc');
      btn.title = collapsed ? 'Expand sidebar' : 'Collapse sidebar';
    });


    document.getElementById('btn-theme').addEventListener('click', () => {
      const isDark = document.documentElement.dataset.theme === 'dark';
      applyTheme(isDark ? 'light' : 'dark');
      try { chrome.storage.local.set({ theme: isDark ? 'light' : 'dark' }); } catch {}
      refreshMermaidTheme();
    });

    // ── Search functionality ───────────────────────────────────────────────────────
    const searchInput = document.getElementById('search-input');
    let searchTimeout;

    searchInput.addEventListener('input', (e) => {
      clearTimeout(searchTimeout);
      searchTimeout = setTimeout(() => {
        const content = document.getElementById('md-content');
        if (content) {
          highlightSearchMatches(e.target.value, content);
        }
      }, 200);
    });

    searchInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        if (e.shiftKey) {
          prevMatch();
        } else {
          nextMatch();
        }
      } else if (e.key === 'Escape') {
        closeSearch();
      }
    });

    document.getElementById('search-close').addEventListener('click', closeSearch);
    document.getElementById('search-next').addEventListener('click', nextMatch);
    document.getElementById('search-prev').addEventListener('click', prevMatch);

    document.getElementById('btn-search').addEventListener('click', () => {
      const searchContainer = document.getElementById('search-container');
      const isHidden = searchContainer.classList.contains('search-hidden');
      if (isHidden) {
        openSearch();
      } else {
        closeSearch();
      }
    });

    // Global keyboard shortcut for search (Ctrl+F / Cmd+F)
    document.addEventListener('keydown', (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 'f') {
        e.preventDefault();
        openSearch();
      }
    });

    // Intercept hash-link clicks — Chrome blocks file:// → file://# navigation.
    // Covers both bare "#section" hrefs and same-file relative links like "doc.md#section".
    document.getElementById('md-root').addEventListener('click', e => {
      const a = e.target.closest('a[href]');
      if (!a) return;
      const attr = a.getAttribute('href') || '';
      let id = null;
      if (attr.startsWith('#')) {
        id = decodeURIComponent(attr.slice(1));
      } else {
        try {
          const resolved = new URL(attr, location.href);
          if (resolved.pathname === location.pathname && resolved.hash) {
            id = decodeURIComponent(resolved.hash.slice(1));
          }
        } catch {}
      }
      if (id === null) return;
      e.preventDefault();
      if (!id) return;
      const el = document.getElementById(id);
      if (el) {
        scrollToHeading(el, 'smooth');
        history.replaceState(null, '', '#' + id);
        const toc = document.getElementById('md-toc');
        if (toc) {
          toc.querySelectorAll('a.active').forEach(a => a.classList.remove('active'));
          const link = toc.querySelector(`a[href="#${CSS.escape(id)}"]`);
          if (link) link.classList.add('active');
        }
      }
    });

    // ── Load saved theme ──────────────────────────────────────────────────────
    try {
      chrome.storage.local.get('theme', data => applyTheme(data.theme || 'auto'));
    } catch { applyTheme('auto'); }
  }

  // ── Scroll Spy ───────────────────────────────────────────────────────────────

  function setupScrollSpy() {
    const toc = document.getElementById('md-toc');
    if (!toc) return;
    const headings = [...document.querySelectorAll('#md-content h1,#md-content h2,#md-content h3,#md-content h4')];
    if (!headings.length) return;
    const obs = new IntersectionObserver(entries => {
      if (scrollSpyLock) return; // click-driven scroll in flight; don't override
      entries.forEach(e => {
        const link = toc.querySelector(`a[href="#${e.target.id}"]`);
        if (!link) return;
        if (e.isIntersecting) {
          toc.querySelectorAll('a.active').forEach(a => a.classList.remove('active'));
          link.classList.add('active');
          // auto-expand any collapsed ancestor sections
          let el = link.parentElement;
          while (el && el !== toc) {
            if (el.tagName === 'LI' && !el.classList.contains('toc-open')) {
              el.classList.add('toc-open');
              const btn = el.querySelector('.toc-toggle');
              if (btn) btn.setAttribute('aria-expanded', 'true');
            }
            el = el.parentElement;
          }
        }
      });
    }, { rootMargin: '-10% 0px -80% 0px' });
    headings.forEach(h => obs.observe(h));
  }

  // ── Init ─────────────────────────────────────────────────────────────────────

  function init() {
    if (!isMarkdownPage()) return;
    history.scrollRestoration = 'manual';
    const pre = document.querySelector('pre');
    const src = pre ? pre.textContent : document.body.innerText;
    if (!src.trim()) return;

    const startHash = decodeURIComponent(location.hash.slice(1));
    if (startHash) history.replaceState(null, '', location.pathname + location.search);

    render(src);

    if (startHash) {
      requestAnimationFrame(() => {
        const el = document.getElementById(startHash);
        if (el) {
          scrollToHeading(el, 'instant');
          history.replaceState(null, '', '#' + startHash);
        }
      });
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

})();
