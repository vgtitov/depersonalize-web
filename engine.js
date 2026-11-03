// Движок обезличивания. Без внешних зависимостей и без сети.
const FIRST = ("Александр Алексей Андрей Антон Артём Артем Борис Вадим Валентин Валерий Василий Виктор Виталий Владимир Владислав Вячеслав Геннадий Георгий Глеб Григорий Даниил Денис Дмитрий Евгений Егор Иван Игорь Илья Кирилл Константин Леонид Максим Михаил Никита Николай Олег Павел Пётр Петр Роман Сергей Станислав Степан Тимофей Фёдор Федор Юрий Ян Яков " +
 "Александра Алёна Алена Алина Алла Анастасия Анна Антонина Валентина Валерия Вера Вероника Виктория Галина Дарья Диана Евгения Екатерина Елена Елизавета Жанна Зоя Инна Ирина Кристина Ксения Лариса Лера Лидия Любовь Людмила Мария Марина Надежда Наталья Наталия Нина Оксана Ольга Полина Раиса Светлана София Софья Татьяна Ульяна Юлия Юлиана").split(" ");
FIRST.push(...("Ярослав Анатолий Руслан Артур Богдан Эдуард Тимур Марк Мирон Назар Назарий Родион Савелий Семён Семен Светозар Ростислав Радим Платон Лев Матвей Кузьма Климент Елисей Ефим Давид Демид Данила Данил Всеволод Вениамин Влад Альберт Аркадий Арсений Артемий Афанасий Бронислав Вальдемар Витольд Вацлав Гавриил Герман Григорий Дамир Денис Захар Зиновий Ильяс Иосиф Казимир Касьян Лазарь Макар Мартин Мстислав Никанор Остап Прохор Ренат Рустам Самуил Тарас Тихон Трофим Устин Фаддей Филипп Харитон Эммануил Эрик Юлиан Юрий Яромир Ярослава Агата Агния Адель Аделина Аксинья Алевтина Альбина Аврора Ангелина Антонида Арина Ася Белла Богдана Вилена Влада Владислава Галя Глафира Дана Дина Доминика Ева Евдокия Елена Есения Зарина Зинаида Злата Изабелла Илона Инга Ия Калерия Каролина Клавдия Клара Лада Лиана Лилия Любава Майя Маргарита Марфа Матрёна Мила Милана Мирослава Муза Нелли Ника Нинель Оля Павла Пелагея Рада Регина Рената Римма Роза Руслана Руфина Сабина Серафима Снежана Стелла Стефания Таисия Тамара Таня Ульяна Устинья Фаина Фёкла Эвелина Эльвира Эмилия Эмма Юлиана Яна Ярослава Эльза Эля Ханна Ольга Лариса").split(" "));
const FIRSTSET = new Set(FIRST.map(s=>s.toLowerCase()));
const STOP = new Set("магазин господин гражданин дворянин крестьянин раввин цин".split(" ")); // слова на -ин/-ов, не фамилии
const COMMON = new Set("марка рада роза злата вера надежда любовь лада мира".split(" ")); // слова, совпадающие с именем
const VERB = /(?:[ао]л|ла|ли|ло|ет|ёт|ит|ут|ют|ат|ят|ть|ться|лся|лась)$/i;
const NOTSURN = new Set("вам вас вами ваш наш январь февраль март апрель май июнь июль август сентябрь октябрь ноябрь декабрь понедельник вторник среда четверг пятница суббота воскресенье вчера сегодня завтра".split(" "));
const CAP = "[А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?";
const PATR = "(?:(?:ович|евич|ьич)(?:а|у|ем|ом|е|ы)?|(?:овн|евн|ичн|инична?)(?:а|у|ы|е|ой|ою)|ична|инична)";
const SURN = "(?:ов|ев|ёв|ин|ын|(?:ск|цк)(?:ий|ая|ого|ому|им|ом|ой|ую|ое)|ова|ева|ина|ына|енко|ук|юк|чик|ич)(?:а|у|ым|ом|ем|ой|е|ых|ого|ому)?";
const SURN1 = "(?:ов|ев|ёв|ын|(?:ск|цк)(?:ий|ая|ого|ому|им|ом|ой|ую|ое)|ова|ева|ина|ына|енко)(?:а|у|ым|ом|ем|ой|е|ых|ого|ому)?"; // для одиночного инициала без «ин»: Магазин, Бензин не фамилии
const ORGF = "(?:СООО|ИООО|ЧТУП|ЧПУП|РУП|КУП|ООО|ЗАО|ОАО|ОДО|ЧУП|ПАО|ЧП|УО|НКО|ГП|ГУ|АО|УП|ИП)";
const ORGW = "(?:[А-ЯЁ][а-яё]{2,}|[A-Z][A-Za-z]{2,}|[А-ЯЁ]{3,}|[A-Z]{3,})(?:-(?:[А-ЯЁ][а-яё]+|[A-Za-z]+|[А-ЯЁ]+))*";
// Название без кавычек: форма собственности и слова после неё, пока не встретилось имя человека.
function orgName(v){const ws=[...v.matchAll(/\S+/g)];let end=ws[0].index+ws[0].length;
  for(let i=1;i<ws.length;i++){const w=ws[i][0],l=w.toLowerCase();
    if(i>1&&(FIRSTSET.has(stem(w))||SURX.test(l)||PATX.test(l)))break;
    end=ws[i].index+w.length}
  return [[0,end]]}
const PATX = /^[а-яё]+(?:(?:ович|евич|ьич)(?:а|у|ем|ом|е|ы)?|(?:овн|евн|ичн|иничн)(?:а|у|ы|е|ой|ою))$/;
const SURX = /(?:ов|ев|ёв|ин|ын|(?:ск|цк)(?:ий|ая|ого|ому|им|ом|ой|ую|ое)|ова|ева|ина|ына|енко|ук|юк)(?:а|у|ым|ом|ем|ой|е|ых|ого|ому)?$/;
// Прописная цепочка «КЛИЕНТ ИВАНОВ ИВАН ИВАНОВИЧ»: находим внутри неё границы ФИО по классам слов (Ф фамилия, И имя, О отчество).
function capsPeople(v){
  const ws=[...v.matchAll(/[А-ЯЁ]{2,}/g)].map(m=>({a:m.index,b:m.index+m[0].length,c:cls(m[0])}));
  const pats=["SFP","PFP","FPS","FP","SF","FS","PF"], out=[];
  for(let i=0;i<ws.length;){
    const p=pats.find(p=>i+p.length<=ws.length&&[...p].every((c,k)=>ws[i+k].c===c));
    if(p){out.push([ws[i].a,ws[i+p.length-1].b]);i+=p.length}else i++}
  return out}
function cls(w){const l=w.toLowerCase();return PATX.test(l)?"P":FIRSTSET.has(stem(w))?"F":SURX.test(l)?"S":"X"}
const QUAL = "(?:[ \\t]+[\\wА-Яа-яЁё-]+){0,9}";
const SEPC = QUAL + "[\"']?[ \\t]*[:=][ \\t]*(?:\\r?\\n|\u2028)?[ \\t]*"; // «пароль от почты: значение», значение может быть на следующей строке
const SEPD = QUAL + "[ \\t]+[-–—][ \\t]+";
const SECVAL = "(?:\"((?:[^\"\\\\\\n\u2028]|\\\\.)+)\"|'((?:[^'\\\\\\n\u2028]|\\\\.)+)'|«([^»\\n\u2028]+)»|((?:[^\\s\"',;&]|&(?!\\w+=)|[,;](?=[^\\s,;]))+?)(?=\\.*(?:[\\s&\"']|[,;](?:\\s|$)|$)))";
const KWS = "парол[ьяю]|password|passwd|passphrase|pass|pwd|api[ _-]?key|apikey|access[ _-]?key|pin|пин|pin-?код|пин-?код|cvv|cvc|otp|код[ ]доступа|код[ ]подтверждения|одноразов\\w+[ ]код|access[ ]code", KWW = "токен|token|секрет|secret";
const notLabel = v => !/^\[[А-ЯЁA-Z]+_\d+\]$/.test(v), strongVal = v => notLabel(v) && (v.length >= 16 || /\d/.test(v) || /[a-zа-яё][A-ZА-ЯЁ]|[A-ZА-ЯЁ].*[a-zа-яё]/.test(v) && /[A-Za-zА-Яа-яЁё]/.test(v) && /[^A-Za-zА-Яа-яЁё]/.test(v));
const kwRe = (kw, sep) => new RegExp(`(?<![\\wА-Яа-яЁё])(?:${kw})(?![\\wА-Яа-яЁё])${sep}${SECVAL}`, "gi");
const RULES = [ // [метка, regex, группа значения, проверка, разбор цепочки]
 ["СЕКРЕТ", kwRe(KWS, SEPC), [1,2,3,4], notLabel],
 ["СЕКРЕТ", kwRe(KWW, SEPC), [1,2,3,4], strongVal],
 ["СЕКРЕТ", kwRe(KWS + "|" + KWW, SEPD), [1,2,3,4], strongVal],
 ["СЕКРЕТ", /(?<![\wА-Яа-яЁё])(?:парол[ьяю]|password|pass|pwd|токен|token|пин|pin)(?:[ \t]+(?:от|для|к|на|в)[ \t]+[\wА-Яа-яЁё-]+(?:[ \t]+[\wА-Яа-яЁё-]+){0,7})?(?:[ \t]+(?:is|есть))?[ \t]+([^\s"',;]{4,})/gi, 1, v=>/\d/.test(v)&&notLabel(v)],
 ["СЕКРЕТ", /(?<![\wА-Яа-яЁё])(?:pin|пин|pin-?код|пин-?код|cvv2?|cvc2?|otp)[ \t]*[:=]?[ \t]*(\d{3,8})(?![\d\wА-Яа-яЁё])/gi, 1],
 ["СЕКРЕТ", kwRe(KWS + "|token|токен", QUAL + "[ \\t]+(?=[\"'«])"), [1,2,3,4], notLabel],
 ["СЕКРЕТ", /(?<=\b[a-z][a-z0-9+.-]*:\/\/[^\s:@\/?#]+:)[^\s@\/?#]+(?=@)/gi, 0],
 ["СЕКРЕТ", /(?<![\wА-Яа-яЁё])bearer[ \t]+([A-Za-z0-9._~+\/=-]{12,})/gi, 1],
 ["СЕКРЕТ", /(?<![\w-])(?:gh[pousr]_[A-Za-z0-9]{30,}|github_pat_[A-Za-z0-9_]{30,}|sk-[A-Za-z0-9_-]{20,}|AKIA[0-9A-Z]{16}|xox[abprs]-[A-Za-z0-9-]{10,}|AIza[0-9A-Za-z_-]{30,}|eyJ[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,})(?![\w-])/g, 0],
 ["ОРГ", new RegExp(`(?<![А-ЯЁа-яёA-Za-z])${ORGF}[ \\t]+[«"„“'][^«»"„“”'\\n\u2028]{1,80}[»"“”']`,"g"), 0],
 ["ОРГ", /(?<![А-ЯЁа-яё])(?:[Оо]бщество с ограниченной ответственностью|[Оо]ткрытое акционерное общество|[Зз]акрытое акционерное общество|[Аа]кционерное общество|[Уу]нитарное предприятие|[Чч]астное унитарное предприятие|[Ии]ндивидуальный предприниматель)[ \t]+[«"„“'][^«»"„“”'\n\u2028]{1,80}[»"“”']/g, 0],
 ["ОРГ", new RegExp(`(?<![А-ЯЁа-яёA-Za-z])(?:СООО|ИООО|ЧТУП|ЧПУП|РУП|КУП|ООО|ЗАО|ОАО|ОДО|ЧУП|ПАО|ЧП|УО|НКО|АО|УП)[ \\t]+${ORGW}(?:[ \\t]+${ORGW}){0,3}(?![А-ЯЁа-яёA-Za-z])`,"g"), 0, null, orgName],
 ["EMAIL", /(?<![\wА-Яа-яЁё.+-])[\wА-Яа-яЁё.+-]{1,64}@[\wА-Яа-яЁё-]{1,63}(?:\.[\wА-Яа-яЁё-]{1,63}){1,8}/g, 0],
 ["ПУТЬ", /file:\/\/\/[^\s"<>]+|(?<![A-Za-z])[A-Za-z]:\\Users\\[^\s"<>|]+/gi, 0],
 ["URL", /\b(?:https?:\/\/|www\.)[^\s<>"')\]]*[^\s<>"')\]\.,;:!?]/gi, 0],
 ["IP", /(?<![\w.])(?:(?:25[0-5]|2[0-4]\d|1?\d?\d)\.){3}(?:25[0-5]|2[0-4]\d|1?\d?\d)(?![\w]|\.\d)/g, 0],
 ["ТЕЛЕФОН", /(?<![\w+])(?:\+375|375|8[\s-]?0)[\s()-]*\d{2}[\s()-]*\d{3}[\s-]*\d{2}[\s-]*\d{2}(?!\d)/g, 0],
 ["ТЕЛЕФОН", /(?<![\w+])(?:\+7|8)[\s()-]*\d{3}[\s()-]*\d{3}[\s-]*\d{2}[\s-]*\d{2}(?!\d)/g, 0],
 ["ТЕЛЕФОН", /(?<![\w+])\+\d{1,3}[\s()-]*\d[\d\s()-]{7,13}\d(?!\d)/g, 0],
 ["ТЕЛЕФОН", /(?<![\w+.:-])\(?(?:80|0)(?:25|29|33|44|17)[\s()-]*\d{3}[\s-]*\d{2}[\s-]*\d{2}(?![\d:])/g, 0],
 ["ТЕЛЕФОН", /(?<![\w+.:()-])(?:25|29|33|44)[ -]\d{3}[ -]\d{2}[ -]\d{2}(?![\d:])/g, 0],
 ["НИК", /(?<![\w/@.])@[A-Za-z]\w{4,31}/g, 0],
 ["ИНН", /(?<![\wА-Яа-яЁё])инн(?![\wА-Яа-яЁё])\D{0,12}?(\d{12}|\d{10})(?!\d)/gi, 1],
 ["УНП", /(?<![\wА-Яа-яЁё])унп(?![\wА-Яа-яЁё])\D{0,12}?(\d{9})(?!\d)/gi, 1],
 ["КАРТА", /(?<!\d)\d(?:[ -]?\d){12,18}(?!\d)/g, 0, v=>luhn(v.replace(/\D/g,""))],
 ["IBAN", /\b[A-Z]{2}\d{2}(?: ?[A-Z0-9]{4}){2,7}(?: ?[A-Z0-9]{1,4})?\b/g, 0],
 ["ЧЕЛОВЕК", /(?<![А-ЯЁа-яё])[А-ЯЁ]{2,}(?:[ \t]+[А-ЯЁ]{2,})+(?![А-ЯЁа-яё])/g, 0, null, capsPeople],
 ["ЧЕЛОВЕК", new RegExp(`(?<![А-ЯЁа-яё])[А-ЯЁ][а-яё]+${SURN}(?![а-яё])[ \\t]+[А-ЯЁ]\\.[ \\t]?[А-ЯЁ]\\.`,"g"), 0],
 ["ЧЕЛОВЕК", new RegExp(`(?<![А-ЯЁа-яё])[А-ЯЁ]\\.[ \\t]?[А-ЯЁ]\\.[ \\t]?[А-ЯЁ][а-яё]+${SURN}(?![а-яё])`,"g"), 0],
 ["ЧЕЛОВЕК", new RegExp(`(?<![А-ЯЁа-яё])[А-ЯЁ][а-яё]+[ \\t]+[А-ЯЁ][а-яё]+${PATR}(?![а-яё])(?:[ \\t]+[А-ЯЁ][а-яё]{2,}(?:-[А-ЯЁ][а-яё]+)?(?![а-яё-]))?`,"g"), 0, null, patrFull],
 ["ЧЕЛОВЕК", new RegExp(`(?<![А-ЯЁа-яё])${CAP}[ \\t]+${CAP}[ \\t]+[А-ЯЁ][а-яё]+${PATR}(?![а-яё])`,"g"), 0, v=>{const [w1,w2]=v.split(/\s+/);return FIRSTSET.has(stem(w2))&&(FIRSTSET.has(stem(w1))||new RegExp("^[А-ЯЁ][а-яё]+"+SURN+"$").test(w1))}],
 ["ЧЕЛОВЕК", new RegExp(`(?<![А-ЯЁа-яё])[А-ЯЁ][а-яё]+${SURN}(?![а-яё])[ \\t]+[А-ЯЁ][а-яё]+(?![а-яё])`,"g"), 0, v=>FIRSTSET.has(stem(v.split(/\s+/)[1]))&&!STOP.has(v.split(/\s+/)[0].toLowerCase())],
 ["ЧЕЛОВЕК", new RegExp(`(?<![А-ЯЁа-яё])[А-ЯЁ][а-яё]+(?![а-яё])[ \\t]+[А-ЯЁ][а-яё]+${SURN}(?![а-яё])`,"g"), 0, v=>FIRSTSET.has(stem(v.split(/\s+/)[0]))],
 ["ЧЕЛОВЕК", /(?<![А-ЯЁа-яё])[А-ЯЁ][а-яё]+(?![а-яё])[ \t]+[А-ЯЁ][а-яё]{2,}(?![а-яё-])/g, 0, v=>{const [a,b]=v.split(/\s+/);return FIRSTSET.has(stem(a))&&!COMMON.has(a.toLowerCase())&&!FIRSTSET.has(stem(b))&&!NOTSURN.has(b.toLowerCase())&&!PATX.test(b.toLowerCase())}],
 ["ЧЕЛОВЕК", new RegExp(`(?<![А-ЯЁа-яё])[А-ЯЁ]\\.[ \\t]?[А-ЯЁ][а-яё]+${SURN1}(?![а-яё])`,"g"), 0],
 ["ЧЕЛОВЕК", new RegExp(`(?<![А-ЯЁа-яё])[А-ЯЁ][а-яё]+${SURN1}(?![а-яё])[ \\t]+[А-ЯЁ]\\.(?![А-ЯЁа-яё])`,"g"), 0],
];
// «Имя Отчество [Фамилия]»: третье слово берём любое с заглавной, но не похожее на глагол
function patrFull(v){const ws=[...v.matchAll(/\S+/g)];if(!FIRSTSET.has(stem(ws[0][0])))return [];
  const last=ws.length>2&&!VERB.test(ws[2][0].split("-")[0])?ws[2]:ws[1];return [[0,last.index+last[0].length]]}
function stem(w){w=w.toLowerCase();if(/^павл(?:а|у|ом|е|ы)?$/.test(w))return "павел";if(FIRSTSET.has(w))return w;for(const e of ["а","у","ом","ем","е","ы","и","ой","ю","я","ей"]){const s=w.endsWith(e)?w.slice(0,-e.length):null;if(s){for(const t of [s,s+"а",s+"я",s+"ь",s+"й"])if(FIRSTSET.has(t))return t}}return w}
function luhn(d){if(d.length<13)return false;let s=0,a=false;for(let i=d.length-1;i>=0;i--){let n=+d[i];if(a){n*=2;if(n>9)n-=9}s+=n;a=!a}return s%10===0}
function esc(s){return s.replace(/[.*+?^${}()|[\]\\]/g,"\\$&")}
// Один человек в разных формах: «Иван Петрович Сидоров», «Сидорову Ивану», «Сидорова» получают одну метку.
const ENDS=["ого","ому","ыми","ым","им","ом","ем","ой","ою","ую","ых","а","у","е","ы","я","ю"];
function surn(w){w=w.toLowerCase().replace(/ё/g,"е").replace(/(ск|цк)(?:ий|ая|ого|ому|им|ой|ую|ом|ое|их)$/,"$1");
  for(const e of ENDS)if(w.length>e.length+3&&w.endsWith(e))return w.slice(0,-e.length);return w}
function parsePerson(v){const p={first:null,patr:null,last:null,ini:""};
  for(const w of v.split(/[\s.]+/).filter(Boolean)){const l=w.toLowerCase();
    if(w.length===1){p.ini+=l;continue}
    const parts=l.split("-");
    if(PATX.test(l))p.patr=l.replace(/(?:ем|ом|ой|ою|а|у|е|ы)$/,"");
    else if(!p.first&&parts.every(x=>FIRSTSET.has(stem(x))))p.first=stem(parts[0]);
    else p.last=parts.map(surn).join("-")}
  return p}
const init=(p,k)=>k===0?(p.first?p.first[0]:p.ini[0]):(p.patr?p.patr[0]:p.ini[1]);
function samePerson(a,b){
  if(a.patr&&b.patr&&a.patr!==b.patr)return false;
  if(a.last&&b.last)return a.last===b.last&&(!a.first||!b.first||a.first===b.first)&&[0,1].every(k=>!init(a,k)||!init(b,k)||init(a,k)===init(b,k));
  if(a.last||b.last){const [x,y]=a.last?[b,a]:[a,b];return !!x.first&&x.first===y.first&&!!x.patr&&x.patr===y.patr}
  return !!a.first&&a.first===b.first&&a.patr===b.patr}
const words=v=>v.split(/\s+/).length;
// Индекс людей: ищем по фамилии или по имени с отчеством, а не перебором всех.
const fp=p=>(p.first||"")+"|"+(p.patr||"");
const qkeys=p=>p.last?["L|"+p.last,...(p.first&&p.patr?["NL|"+fp(p)]:[])]:p.first?["FP|"+fp(p)]:[];

const SUF="(?:ами|ями|ого|ому|ой|ей|ою|ею|ью|ом|ем|ём|ым|им|ов|ев|ам|ям|ах|ях|а|я|у|ю|ы|и|е|й|ь)?"; // падежные окончания слова из словаря
// Разбор одного текста: непересекающиеся находки [от, до, метка, значение] и секреты внутри них.
function analyze(text, dictRes){
  const spans=[], sec=[];
  for(const [re,w] of dictRes){re.lastIndex=0;let m;while((m=re.exec(text)))spans.push([m.index,m.index+m[0].length,"СЛОВАРЬ",w])}
  for(const [label,re,grp,check,multi] of RULES){re.lastIndex=0;let m;
    while((m=re.exec(text))){const v=Array.isArray(grp)?grp.map(g=>m[g]).find(x=>x!=null):m[grp];if(v==null)continue;
      const st=m.index+(grp?m[0].lastIndexOf(v):0);
      if(multi){for(const [a,b] of multi(v))spans.push([st+a,st+b,label,text.slice(st+a,st+b)]);continue}
      if(check&&!check(v))continue;
      let a=st,b=st+v.length;
      if(label==="ЧЕЛОВЕК"){let h;while((h=/^-[А-ЯЁ][а-яё]+/.exec(text.slice(b,b+60))))b+=h[0].length;
        while((h=/([А-ЯЁ][а-яё]+)-$/.exec(text.slice(Math.max(0,a-60),a))))a-=h[0].length}
      spans.push([a,b,label,text.slice(a,b)]);if(label==="СЕКРЕТ")sec.push([st,st+v.length,v]);}}
  spans.sort((a,b)=>a[0]-b[0]||(b[1]-b[0])-(a[1]-a[0])||(b[2]==="СЕКРЕТ")-(a[2]==="СЕКРЕТ"));
  const keep=[];for(const s of spans){const l=keep[keep.length-1];
    if(l&&s[0]<l[1]){if(s[1]>l[1]){if(s[1]-s[0]>l[1]-l[0]){l[2]=s[2]}l[1]=s[1];l[3]=text.slice(l[0],l[1]);if(l[2]==="СЛОВАРЬ"){l[2]="СЛОВО"}}}
    else keep.push(s.slice())}
  sec.sort((x,y)=>x[0]-y[0]);
  return {keep,sec};
}
// Много текстов (ячейки, абзацы) с общей нумерацией меток и общей памятью о людях.
function depersonalizeMany(texts, dictLines){
  const map=new Map(), byKey=new Map(), counts={}, num={}, idx=new Map();
  const used=new Set();for(const t of texts)for(const m of t.match(/\[[А-ЯЁA-Z]+_\d+\]/g)||[])used.add(m); // номера, уже занятые метками в исходном тексте, пропускаем
  const fresh=(label,val)=>{let t;do{num[label]=(num[label]||0)+1;t=`[${label}_${num[label]}]`}while(used.has(t));
    counts[label]=(counts[label]||0)+1;if(label!=="СЕКРЕТ")map.set(t,val);return t}; // секреты в карту не пишем
  const add=e=>{const p=e.p,ks=[];if(p.last)ks.push("L|"+p.last);if(p.first){ks.push("FP|"+fp(p));if(!p.last)ks.push("NL|"+fp(p))}
    for(const k of ks){if(!idx.has(k))idx.set(k,new Set());idx.get(k).add(e)}};
  const token=(label,val)=>{
    if(label==="ЧЕЛОВЕК"){const p=parsePerson(val);let e=null;
      for(const k of qkeys(p)){for(const c of idx.get(k)||[])if(samePerson(c.p,p)){e=c;break}if(e)break}
      if(e){if(words(val)>words(e.val)){e.val=val;map.set(e.tok,val)}
        e.p.first=e.p.first||p.first;e.p.last=e.p.last||p.last;e.p.patr=e.p.patr||p.patr;e.p.ini=e.p.ini||p.ini;add(e);return e.tok}
      const tok=fresh(label,val);add({p,val,tok});return tok}
    const key=label+"|"+(label==="URL"?val:val.toLowerCase().replace(/\s+/g," "));
    if(byKey.has(key))return byKey.get(key);
    const t=fresh(label,val);byKey.set(key,t);return t};
  // словарь пользователя: слово и его склонения (основа без последней гласной/й/ь), без учёта регистра
  const dictRes=[];for(const raw of dictLines){const w=raw.trim();if(w.length<3)continue;
    const cyr=/[а-яё]/i.test(w), base=esc(cyr&&/[аяьй]$/i.test(w)?w.slice(0,-1):w);
    dictRes.push([new RegExp(`(?<![А-ЯЁа-яёA-Za-z])${base}${cyr?SUF:"(?:'s|s|es)?"}(?![А-ЯЁа-яёA-Za-z])`,"gi"),w])}
  const parsed=texts.map(t=>analyze(t,dictRes));
  // одинокая фамилия, которая в этих текстах названа с именем: «сказала Сидорову»
  const lasts=new Set();
  for(const {keep} of parsed)for(const k of keep)if(k[2]==="ЧЕЛОВЕК"){const l=parsePerson(k[3]).last;if(l&&l.length>=4&&/(?:ов|ев|ёв|ин|ын|ск\w{0,2}|цк\w{0,2}|ова|ева|ина|ына|енко|ко|ло|ук|юк|ич|ыч|ий|ой|ая|ач|ик|ец|ець|аль|ель|оль|арь|ярь|ар|ай|ун|ак|як|ан|ян|ей|ур|аш|ёк|ек)$/i.test(l))lasts.add(l)}
  const outs=texts.map((text,ti)=>{
    const {keep,sec}=parsed[ti];
    if(lasts.size){const re=/(?<![А-ЯЁа-яё])[А-ЯЁ][а-яё]{3,}(?![А-ЯЁа-яё])/g,extra=[];let m,j=0;
      while((m=re.exec(text))){const a=m.index,b=a+m[0].length;
        while(j<keep.length&&keep[j][1]<=a)j++; // keep отсортирован и без наложений
        if(lasts.has(surn(m[0]))&&!FIRSTSET.has(stem(m[0]))&&!(j<keep.length&&keep[j][0]<b))extra.push([a,b,"ЧЕЛОВЕК",m[0]])}
      if(extra.length){keep.push(...extra);keep.sort((a,b)=>a[0]-b[0])}}
    let out="",pos=0,sp=0;
    for(const [a,b,label,val] of keep){out+=text.slice(pos,a);
      let raw=text.slice(a,b);
      if(label!=="СЕКРЕТ"){ // секрет внутри ссылки или другого значения: в карту уходит значение уже с меткой секрета
        while(sp<sec.length&&sec[sp][1]<=a)sp++;
        const inner=[];for(let k=sp;k<sec.length&&sec[k][0]<b;k++)if(sec[k][0]>=a&&sec[k][1]<=b)inner.push(sec[k]);
        const ts=inner.map(([,,v])=>token("СЕКРЕТ",v));
        for(let k=inner.length-1;k>=0;k--)raw=raw.slice(0,inner[k][0]-a)+ts[k]+raw.slice(inner[k][1]-a)}
      out+=token(label==="СЛОВАРЬ"?"СЛОВО":label, label==="СЛОВАРЬ"?val:raw);pos=b} // слово из словаря: одно слово = один токен независимо от склонения
    return out+text.slice(pos)});
  const warn=[];
  if(used.size)warn.push("в исходном тексте уже есть метки вида [ЧЕЛОВЕК_1]: при возврате имён они остаются как есть");
  if(!dictLines.some(s=>s.trim()))warn.push("словарь пуст: фамилии без характерного окончания и названия компаний без формы собственности могли остаться");
  return {texts:outs,map,counts,warn};
}
function depersonalize(text, dictLines){const r=depersonalizeMany([text],dictLines);return {text:r.texts[0],map:r.map,counts:r.counts,warn:r.warn}}
function mapToJson(map){return JSON.stringify({format:"depers-web-map-1",pairs:[...map.entries()]})}
function mapFromJson(j){const o=JSON.parse(j);if(o.format!=="depers-web-map-1"||!Array.isArray(o.pairs))throw new Error("не карта");return new Map(o.pairs.filter(p=>Array.isArray(p)&&typeof p[0]==="string"&&typeof p[1]==="string"))}
function restore(text,map){return text.replace(/\[[А-ЯЁA-Z]+_\d+\]/g,t=>map.has(t)?map.get(t):t)}
if (typeof module !== "undefined") module.exports = { depersonalize, depersonalizeMany, restore, mapToJson, mapFromJson };
