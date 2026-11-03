// Движок обезличивания. Без внешних зависимостей и без сети.
const FIRST = ("Александр Алексей Андрей Антон Артём Артем Борис Вадим Валентин Валерий Василий Виктор Виталий Владимир Владислав Вячеслав Геннадий Георгий Глеб Григорий Даниил Денис Дмитрий Евгений Егор Иван Игорь Илья Кирилл Константин Леонид Максим Михаил Никита Николай Олег Павел Пётр Петр Роман Сергей Станислав Степан Тимофей Фёдор Федор Юрий Ян Яков " +
 "Александра Алёна Алена Алина Алла Анастасия Анна Антонина Валентина Валерия Вера Вероника Виктория Галина Дарья Диана Евгения Екатерина Елена Елизавета Жанна Зоя Инна Ирина Кристина Ксения Лариса Лера Лидия Любовь Людмила Мария Марина Надежда Наталья Наталия Нина Оксана Ольга Полина Раиса Светлана София Софья Татьяна Ульяна Юлия Юлиана").split(" ");
const FIRSTSET = new Set(FIRST.map(s=>s.toLowerCase()));
const CAP = "[А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?";
const PATR = "(?:(?:ович|евич|ьич)(?:а|у|ем|ом|е|ы)?|(?:овн|евн|ичн|инична?)(?:а|у|ы|е|ой|ою)|ична|инична)";
const SURN = "(?:ов|ев|ёв|ин|ын|(?:ск|цк)(?:ий|ая|ого|ому|им|ом|ой|ую|ое)|ова|ева|ина|ына|енко|ук|юк|чик|ич)(?:а|у|ым|ом|ем|ой|е|ых|ого|ому)?";
const ORGF = "(?:ООО|ЗАО|ОАО|ОДО|ЧУП|УП|ПАО|АО|ИП|ЧП|НКО|УО)";
const PATX = /^[а-яё]+(?:(?:ович|евич|ьич)(?:а|у|ем|ом|е|ы)?|(?:овн|евн|ичн|иничн)(?:а|у|ы|е|ой|ою))$/;
const SURX = /(?:ов|ев|ёв|ин|ын|(?:ск|цк)(?:ий|ая|ого|ому|им|ом|ой|ую|ое)|ова|ева|ина|ына|енко|ук|юк)(?:а|у|ым|ом|ем|ой|е|ых|ого|ому)?$/;
// Прописная цепочка «КЛИЕНТ ИВАНОВ ИВАН ИВАНОВИЧ»: находим внутри неё границы ФИО по классам слов (Ф фамилия, И имя, О отчество).
function capsPeople(v){
  const ws=[...v.matchAll(/[А-ЯЁ]{2,}/g)].map(m=>({a:m.index,b:m.index+m[0].length,c:cls(m[0])}));
  const pats=["SFP","FPS","FP","SF","FS"], out=[];
  for(let i=0;i<ws.length;){
    const p=pats.find(p=>i+p.length<=ws.length&&[...p].every((c,k)=>ws[i+k].c===c));
    if(p){out.push([ws[i].a,ws[i+p.length-1].b]);i+=p.length}else i++}
  return out}
function cls(w){const l=w.toLowerCase();return PATX.test(l)?"P":FIRSTSET.has(stem(w))?"F":SURX.test(l)?"S":"X"}
const SEP = "[\"']?(?:[ \\t]*[:=]|[ \\t]+[-–—])[ \\t]*";
const SECVAL = "(?:\"([^\"\\n]{1,200})\"|'([^'\\n]{1,200})'|«([^»\\n]{1,200})»|([^\\s\"',;]+))";
const RULES = [ // [метка, regex, группа значения, проверка, разбор цепочки]
 ["СЕКРЕТ", new RegExp(`(?<![\\wА-Яа-яЁё])(?:парол[ьяю]|password|passwd|passphrase|pass|pwd|токен|token|секрет|secret|api[ _-]?key|apikey|access[ _-]?key)(?![\\wА-Яа-яЁё])${SEP}${SECVAL}`,"gi"), [1,2,3,4], v=>!/^\[[А-ЯЁA-Z]+_\d+\]$/.test(v)],
 ["СЕКРЕТ", /(?<![\wА-Яа-яЁё])bearer[ \t]+([A-Za-z0-9._~+\/=-]{12,})/gi, 1],
 ["СЕКРЕТ", /(?<![\w-])(?:gh[pousr]_[A-Za-z0-9]{30,}|github_pat_[A-Za-z0-9_]{30,}|sk-[A-Za-z0-9_-]{20,}|AKIA[0-9A-Z]{16}|xox[abprs]-[A-Za-z0-9-]{10,}|AIza[0-9A-Za-z_-]{30,}|eyJ[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,})(?![\w-])/g, 0],
 ["ОРГ", new RegExp(`(?<![А-ЯЁа-яёA-Za-z])${ORGF}[ \\t]+[«"„“'][^«»"„“”'\\n]{1,80}[»"“”']`,"g"), 0],
 ["ОРГ", /(?<![А-ЯЁа-яёA-Za-z])(?:ООО|ЗАО|ОАО|ОДО|ЧУП|УП|ПАО|АО)[ \t]+(?:[А-ЯЁ][а-яё]{2,}|[A-Z][A-Za-z]{2,}|[А-ЯЁ]{3,}|[A-Z]{3,})(?:-(?:[А-ЯЁ][а-яё]+|[A-Za-z]+|[А-ЯЁ]+))*(?![А-ЯЁа-яёA-Za-z])/g, 0],
 ["EMAIL", /(?<![\wА-Яа-яЁё.+-])[\wА-Яа-яЁё.+-]{1,64}@[\wА-Яа-яЁё-]{1,63}(?:\.[\wА-Яа-яЁё-]{1,63}){1,8}/g, 0],
 ["URL", /\b(?:https?:\/\/|www\.)[^\s<>"')\]]*[^\s<>"')\]\.,;:!?]/gi, 0],
 ["IP", /(?<![\w.])(?:(?:25[0-5]|2[0-4]\d|1?\d?\d)\.){3}(?:25[0-5]|2[0-4]\d|1?\d?\d)(?![\w]|\.\d)/g, 0],
 ["ТЕЛЕФОН", /(?<![\w+])(?:\+375|375|8[\s-]?0)[\s()-]*\d{2}[\s()-]*\d{3}[\s-]*\d{2}[\s-]*\d{2}(?!\d)/g, 0],
 ["ТЕЛЕФОН", /(?<![\w+])(?:\+7|8)[\s()-]*\d{3}[\s()-]*\d{3}[\s-]*\d{2}[\s-]*\d{2}(?!\d)/g, 0],
 ["ТЕЛЕФОН", /(?<![\w+])\+\d{1,3}[\s()-]*\d[\d\s()-]{7,13}\d(?!\d)/g, 0],
 ["ТЕЛЕФОН", /(?<![\w+.:-])(?:80|0)(?:25|29|33|44|17)[\s()-]*\d{3}[\s-]*\d{2}[\s-]*\d{2}(?![\d:])/g, 0],
 ["НИК", /(?<![\w/@.])@[A-Za-z]\w{4,31}/g, 0],
 ["ИНН", /(?<![\wА-Яа-яЁё])инн(?![\wА-Яа-яЁё])\D{0,12}?(\d{12}|\d{10})(?!\d)/gi, 1],
 ["УНП", /(?<![\wА-Яа-яЁё])унп(?![\wА-Яа-яЁё])\D{0,12}?(\d{9})(?!\d)/gi, 1],
 ["КАРТА", /(?<!\d)\d(?:[ -]?\d){12,18}(?!\d)/g, 0, v=>luhn(v.replace(/\D/g,""))],
 ["IBAN", /\b[A-Z]{2}\d{2}(?: ?[A-Z0-9]{4}){2,7}(?: ?[A-Z0-9]{1,4})?\b/g, 0],
 ["ЧЕЛОВЕК", /(?<![А-ЯЁа-яё])[А-ЯЁ]{2,}(?:[ \t]+[А-ЯЁ]{2,})+(?![А-ЯЁа-яё])/g, 0, null, capsPeople],
 ["ЧЕЛОВЕК", new RegExp(`(?<![А-ЯЁа-яё])[А-ЯЁ][а-яё]+${SURN}(?![а-яё])[ \\t]+[А-ЯЁ]\\.[ \\t]?[А-ЯЁ]\\.`,"g"), 0],
 ["ЧЕЛОВЕК", new RegExp(`(?<![А-ЯЁа-яё])[А-ЯЁ]\\.[ \\t]?[А-ЯЁ]\\.[ \\t]?[А-ЯЁ][а-яё]+${SURN}(?![а-яё])`,"g"), 0],
 ["ЧЕЛОВЕК", new RegExp(`(?<![А-ЯЁа-яё])[А-ЯЁ][а-яё]+[ \\t]+[А-ЯЁ][а-яё]+${PATR}(?![а-яё])(?:[ \\t]+[А-ЯЁ][а-яё]+${SURN}(?![а-яё]))?`,"g"), 0, v=>FIRSTSET.has(stem(v.split(/\s+/)[0]))],
 ["ЧЕЛОВЕК", new RegExp(`(?<![А-ЯЁа-яё])${CAP}[ \\t]+${CAP}[ \\t]+[А-ЯЁ][а-яё]+${PATR}(?![а-яё])`,"g"), 0, v=>{const [w1,w2]=v.split(/\s+/);return FIRSTSET.has(stem(w2))&&(FIRSTSET.has(stem(w1))||new RegExp("^[А-ЯЁ][а-яё]+"+SURN+"$").test(w1))}],
 ["ЧЕЛОВЕК", new RegExp(`(?<![А-ЯЁа-яё])[А-ЯЁ][а-яё]+${SURN}(?![а-яё])[ \\t]+[А-ЯЁ][а-яё]+(?![а-яё])`,"g"), 0, v=>FIRSTSET.has(stem(v.split(/\s+/)[1]))],
 ["ЧЕЛОВЕК", new RegExp(`(?<![А-ЯЁа-яё])[А-ЯЁ][а-яё]+(?![а-яё])[ \\t]+[А-ЯЁ][а-яё]+${SURN}(?![а-яё])`,"g"), 0, v=>FIRSTSET.has(stem(v.split(/\s+/)[0]))],
];
function stem(w){w=w.toLowerCase();if(/^павл(?:а|у|ом|е|ы)?$/.test(w))return "павел";if(FIRSTSET.has(w))return w;for(const e of ["а","у","ом","ем","е","ы","и","ой","ю","я","ей"]){const s=w.endsWith(e)?w.slice(0,-e.length):null;if(s){for(const t of [s,s+"а",s+"я",s+"ь",s+"й"])if(FIRSTSET.has(t))return t}}return w}
function luhn(d){if(d.length<13)return false;let s=0,a=false;for(let i=d.length-1;i>=0;i--){let n=+d[i];if(a){n*=2;if(n>9)n-=9}s+=n;a=!a}return s%10===0}
function esc(s){return s.replace(/[.*+?^${}()|[\]\\]/g,"\\$&")}
// Один человек в разных формах: «Иван Петрович Сидоров», «Сидорову Ивану», «Сидорова» получают одну метку.
const ENDS=["ого","ому","ыми","ым","им","ом","ем","ой","ою","ую","ых","а","у","е","ы","я","ю"];
function surn(w){w=w.toLowerCase().replace(/ё/g,"е").replace(/(ск|цк)(?:ий|ая|ого|ому|им|ой|ую|ом|ое|их)$/,"$1");
  for(const e of ENDS)if(w.length>e.length+3&&w.endsWith(e))return w.slice(0,-e.length);return w}
function parsePerson(v){const p={first:null,patr:null,last:null};
  for(const w of v.split(/[\s.]+/).filter(x=>x.length>1)){const l=w.toLowerCase();
    if(PATX.test(l))p.patr=l.replace(/(?:ем|ом|ой|ою|а|у|е|ы)$/,"");else if(!p.first&&FIRSTSET.has(stem(w)))p.first=stem(w);else p.last=surn(w)}
  return p}
function samePerson(a,b){
  if(a.patr&&b.patr&&a.patr!==b.patr)return false;
  if(a.last&&b.last)return a.last===b.last&&(!a.first||!b.first||a.first===b.first);
  if(a.last||b.last){const [x,y]=a.last?[b,a]:[a,b];return !!x.first&&x.first===y.first&&!!x.patr&&x.patr===y.patr}
  return !!a.first&&a.first===b.first&&a.patr===b.patr}
const words=v=>v.split(/\s+/).length;
const SUF="(?:ами|ями|ого|ому|ой|ей|ою|ею|ью|ом|ем|ём|ым|им|ов|ев|ам|ям|ах|ях|а|я|у|ю|ы|и|е|й|ь)?"; // падежные окончания слова из словаря
function depersonalize(text, dictLines){
  const map=new Map(), byKey=new Map(), counts={}, num={}, spans=[], people=[], sec=[];
  const used=new Set(text.match(/\[[А-ЯЁA-Z]+_\d+\]/g)||[]); // номера, уже занятые метками в исходном тексте, пропускаем
  const fresh=(label,val)=>{let t;do{num[label]=(num[label]||0)+1;t=`[${label}_${num[label]}]`}while(used.has(t));
    counts[label]=(counts[label]||0)+1;if(label!=="СЕКРЕТ")map.set(t,val);return t}; // секреты в карту не пишем
  const token=(label,val)=>{
    if(label==="ЧЕЛОВЕК"){const p=parsePerson(val),e=people.find(e=>samePerson(e.p,p));
      if(e){if(words(val)>words(e.val)){e.val=val;map.set(e.tok,val)}
        e.p.first=e.p.first||p.first;e.p.last=e.p.last||p.last;e.p.patr=e.p.patr||p.patr;return e.tok}
      const tok=fresh(label,val);people.push({p,val,tok});return tok}
    const key=label+"|"+(label==="URL"?val:val.toLowerCase().replace(/\s+/g," "));
    if(byKey.has(key))return byKey.get(key);
    const t=fresh(label,val);byKey.set(key,t);return t};
  // 1) словарь пользователя: слово и его склонения (основа без последней гласной/й/ь), без учёта регистра
  for(const raw of dictLines){const w=raw.trim();if(w.length<3)continue;
    const cyr=/[а-яё]/i.test(w), base=esc(cyr&&/[аяьй]$/i.test(w)?w.slice(0,-1):w);
    const re=new RegExp(`(?<![А-ЯЁа-яёA-Za-z])${base}${cyr?SUF:"[a-z]{0,2}"}(?![А-ЯЁа-яёA-Za-z])`,"gi");
    let m;while((m=re.exec(text)))spans.push([m.index,m.index+m[0].length,"СЛОВАРЬ",w]);}
  // 2) правила
  for(const [label,re,grp,check,multi] of RULES){re.lastIndex=0;let m;
    while((m=re.exec(text))){const v=Array.isArray(grp)?grp.map(g=>m[g]).find(x=>x!=null):m[grp];if(v==null)continue;
      const st=m.index+(grp?m[0].lastIndexOf(v):0);
      if(multi){for(const [a,b] of multi(v))spans.push([st+a,st+b,label,text.slice(st+a,st+b)]);continue}
      if(check&&!check(v))continue;spans.push([st,st+v.length,label,v]);if(label==="СЕКРЕТ")sec.push([st,st+v.length,v]);}}
  // 3) без наложений: длинные раньше
  spans.sort((a,b)=>a[0]-b[0]||(b[1]-b[0])-(a[1]-a[0]));
  const keep=[];for(const s of spans){const l=keep[keep.length-1];
    if(l&&s[0]<l[1]){if(s[1]>l[1]){if(s[1]-s[0]>l[1]-l[0]){l[2]=s[2]}l[1]=s[1];l[3]=text.slice(l[0],l[1]);if(l[2]==="СЛОВАРЬ"){l[2]="СЛОВО"}}}
    else keep.push(s.slice())}
  // 4) одинокая фамилия, которая в этом же тексте названа с именем: «сказала Сидорову»
  const lasts=new Set(keep.filter(k=>k[2]==="ЧЕЛОВЕК").map(k=>parsePerson(k[3]).last).filter(l=>l&&l.length>=4));
  if(lasts.size){const re=/(?<![А-ЯЁа-яё])[А-ЯЁ][а-яё]{3,}(?![А-ЯЁа-яё])/g,extra=[];let m,j=0;
    while((m=re.exec(text))){const a=m.index,b=a+m[0].length;
      while(j<keep.length&&keep[j][1]<=a)j++; // keep отсортирован и без наложений
      if(lasts.has(surn(m[0]))&&!FIRSTSET.has(stem(m[0]))&&!(j<keep.length&&keep[j][0]<b))extra.push([a,b,"ЧЕЛОВЕК",m[0]])}
    keep.push(...extra);keep.sort((a,b)=>a[0]-b[0])}
  sec.sort((x,y)=>x[0]-y[0]);
  let out="",pos=0,sp=0;
  for(const [a,b,label,val] of keep){out+=text.slice(pos,a);
    let raw=text.slice(a,b);
    if(label!=="СЕКРЕТ"){ // секрет внутри ссылки или другого значения: в карту уходит значение уже с меткой секрета
      while(sp<sec.length&&sec[sp][1]<=a)sp++;
      const inner=[];for(let k=sp;k<sec.length&&sec[k][0]<b;k++)if(sec[k][0]>=a&&sec[k][1]<=b)inner.push(sec[k]);
      const ts=inner.map(([,,v])=>token("СЕКРЕТ",v));
      for(let k=inner.length-1;k>=0;k--)raw=raw.slice(0,inner[k][0]-a)+ts[k]+raw.slice(inner[k][1]-a)}
    // имя из словаря: одно слово = один токен независимо от склонения
    out+=token(label==="СЛОВАРЬ"?"СЛОВО":label, label==="СЛОВАРЬ"?val:raw);pos=b}
  out+=text.slice(pos);
  const warn=[];
  if(/\[[А-ЯЁA-Z]+_\d+\]/.test(text))warn.push("в исходном тексте уже есть метки вида [ЧЕЛОВЕК_1]: при возврате имён они остаются как есть");
  if(!dictLines.some(s=>s.trim()))warn.push("словарь пуст: фамилии без характерного окончания и названия компаний без формы собственности могли остаться");
  return {text:out,map,counts,warn};
}
function mapToJson(map){return JSON.stringify({format:"depers-web-map-1",pairs:[...map.entries()]})}
function mapFromJson(j){const o=JSON.parse(j);if(o.format!=="depers-web-map-1"||!Array.isArray(o.pairs))throw new Error("не карта");return new Map(o.pairs.filter(p=>Array.isArray(p)&&typeof p[0]==="string"&&typeof p[1]==="string"))}
function restore(text,map){return text.replace(/\[[А-ЯЁA-Z]+_\d+\]/g,t=>map.has(t)?map.get(t):t)}
if (typeof module !== "undefined") module.exports = { depersonalize, restore, mapToJson, mapFromJson };
