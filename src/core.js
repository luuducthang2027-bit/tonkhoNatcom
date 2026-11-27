const NT={};(function(){
const NV={mi:'Hàng mới/New',noc:'UCTT/NOC',th:'Hàng thu hồi',kcs:'Chờ KCS/Waitting for KCS',tl:'Thanh lý/Liquiment',un:'Chưa phân loại'};
const BR=['ART','CEN','NOR','NOT','OUE01','OUE02','OUE03','SUD'],DAY=864e5,p2=n=>('0'+n).slice(-2);
const nm=v=>String(v==null?'':v).replace(/\s+/g,' ').trim().toLowerCase();
const col=(m,p)=>{const k=Object.keys(m).find(k=>k.startsWith(p));return k===undefined?-1:m[k]};
const dt=v=>{if(v instanceof Date){if(isNaN(v))return null;const x=new Date(v.getTime()+432e5);x.setHours(0,0,0,0);return x}
 if(typeof v==='number'&&v>30000)return new Date(Math.round((v-25569)*DAY)+432e5-432e5);
 const s=String(v);return/^\d{1,2}\/\d{1,2}\/\d{4}/.test(s)?new Date(+s.slice(6,10),+s.slice(3,5)-1,+s.slice(0,2)):null};
const mk=d=>d.getFullYear()+'-'+p2(d.getMonth()+1);
/* đọc 1 file: tìm sheet + dòng tiêu đề chứa đủ các cột cần */
const NEED={ct:['code warehouse','code list'],tt:['warehouse code','amount'],td:['shipment code'],sn:['serial number','code of stock-in request'],tl:['warehouse export','cash total'],ex:['stock out','cash total'],im:['stock import','cash total']};NT.NEED=NEED;
NT.LBL={cat:'Danh mục Loai Kho + Nhom VTTB',ct:'Tồn biểu đồ',tt:'Tồn tiền',td:'Tiến độ lô hàng',ncc:'Nhập NCC',pnk:'Nhập thu hồi',sn:'Thu hồi chi tiết (serial)',ex:'Xuất kho',tl:'Xuất thanh lý','?':'Không nhận ra'};
NT.find=(wb,need)=>{for(const sn of wb.SheetNames){const a=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,raw:true,defval:''});
  for(let i=0;i<Math.min(15,a.length);i++){const m={};[a[i],a[i+1]||[]].forEach(row=>row.forEach((v,j)=>{const k=nm(v);if(k&&!(k in m))m[k]=j}));
   if(need.every(n=>n in m)&&need.some(n=>a[i].some(v=>nm(v)===n)))return{m,rows:a.slice(i+1).filter(r=>r.some(x=>x!==''))}}}return null};
NT.sub=d=>{const g=n=>col(d.m,n),a=g('supplier name'),b=g('withdraw from unit'),c=g('reason');let x=0,y=0;d.rows.forEach(r=>{if(String(r[a]||'').trim())x++;if(String(r[b]||'').trim()||/recall|thu hồi/i.test(String(r[c]||'')))y++});return x>=y?'ncc':'pnk'};
NT.detect=(buf,name)=>{const ns=XLSX.read(buf,{type:'array',bookSheets:true}).SheetNames.map(nm);
 if(ns.includes('loai kho')&&ns.includes('nhom vttb'))return{kind:'cat',why:'có sheet Loai Kho + Nhom VTTB',cat:NT.cat(buf)};
 const wb=XLSX.read(buf,{type:'array',cellDates:true});
 for(const k of['ct','tt','sn','td','tl','ex','im']){const d=NT.find(wb,NEED[k]);if(d){let kind=k,why='tiêu đề cột: '+NEED[k].join(' + ');
  if(k==='im'){kind=/thu.?h[oồ]i|recall/i.test(name)?'pnk':/ncc|supplier/i.test(name)?'ncc':NT.sub(d);why+=kind==='ncc'?' · có nhà cung cấp/hợp đồng':' · có đơn vị/lý do thu hồi'}return{kind,why,wb,d}}}return null};
const sk=v=>{if(v==null||v==='')return'';let s=typeof v==='number'?(Math.abs(v)>=1e15?v.toLocaleString('fullwide',{useGrouping:false}):String(parseFloat(v.toPrecision(15)))):String(v).trim();if(/^\d{16,}$/.test(s))s=s.slice(0,15)+'~'+s.length;return s.toLowerCase()};
/* danh mục Loai Kho + Nhom VTTB từ workbook tháng trước */
NT.cat=buf=>{const wb=XLSX.read(buf,{type:'array',sheets:['Loai Kho','Nhom VTTB']}),lk={},nh={};
 XLSX.utils.sheet_to_json(wb.Sheets['Loai Kho'],{header:1,defval:''}).slice(1).forEach(r=>{if(r[1]&&r[2])lk[String(r[1]).trim()]=String(r[2]).trim()});
 XLSX.utils.sheet_to_json(wb.Sheets['Nhom VTTB'],{header:1,defval:''}).slice(2).forEach(r=>{if(r[1])nh[String(r[1]).trim()]={g:String(r[3]).trim(),t:String(r[4]).trim()}});
 return{lk,nh}};
const P=(rs,rf,cf,vf)=>{const m=new Map(),cs=new Set();rs.forEach(r=>{const a=rf(r),b=cf(r);cs.add(b);const x=m.get(a)||{};x[b]=(x[b]||0)+vf(r);m.set(a,x)});return{rk:[...m.keys()].sort(),ck:[...cs].sort(),m}};
const A=(t,p)=>{const rows=p.rk.map(k=>{const x=p.m.get(k),s=p.ck.map(c=>x[c]||0);return[k,...s,s.reduce((a,b)=>a+b,0)]});
 return[[t,...p.ck,'Tổng'],...rows,['Tổng',...(rows[0]||[]).slice(1).map((_,i)=>rows.reduce((a,r)=>a+r[i+1],0))]]};
/* ===== quy tắc nhận dạng loại thiết bị, hãng: lấy từ tool mẫu VTG V15.19 ===== */
function fold(s) { return String(s === null || s === undefined ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').replace(/\s+/g, ' ').trim().toLowerCase(); }
var ACC = /(^|[^a-z])(cap|cable|day|wire|kit|ga|kep|clamp|connector|chau|bulong|bolt|ban|plate|jumper|op|veneer|moc|hook|phu kien|accessor|parts?|ret|aisg|mounting|installation|bracket|screw|label|tie|ong|pipe|tem|sticker|foundation|dai|strap|gioang|seal|nap|cover|tilt|feeder|coaxial|patch)([^a-z]|$)/;
  var ACCW = /^(cap|cable|day|wire|kit|ga|kep|clamp|connector|chau|bulong|bolt|ban|plate|jumper|op|veneer|moc|hook|phu kien|accessor|parts?|ret|aisg|mounting|installation|bracket|screw|label|tie|ong|pipe|tem|sticker|foundation|dai|strap|gioang|seal|nap|cover|tilt|feeder|coaxial|patch|power|nguon|dau cot|jointing)$/;
  function accH(f) { return f.split('|').some(function (part) { var w = part.trim().split(/[\s,()\-]+/).slice(0, 3); return w.some(function (x, i) { return ACCW.test(x) || (i < 2 && ACCW.test(w.slice(i, i + 2).join(' '))); }); }); }
  var CATS = [
    ['fcab', 'Cáp quang', 'km', function (f) { return /cap quang|optical (fiber |fibre )?cable|fib(er|re) (optic )?cable|\badss\b|\bofc\b|\bfo cable|cable fo\b/.test(f) && !/mang xong|closure|odf|kep|clamp|phu kien|hop noi|splice|pigtail|day nhay|patch|connector|ban op|veneer|mo rong|\bdlc\b|\bupc\b|\bapc\b|\blc ?\/|\bsc ?\/|cable parts/.test(f); }, 'ngv'],
    ['polec', 'Cột gỗ / cột treo cáp', 'cột', function (f) { return /cot go|wood(en)? pole|cot be tong|concrete pole|cot dien|cot treo cap|cot thep tron|steel pole|\bpole\b/.test(f) && !/ban op|bulong|bolt|moc|veneer|hook|plate|ban dinh vi|ga |kep|clamp|bracket/.test(f); }, 'ngv'],
    ['ngacc', 'Phụ kiện ngoại vi (măng xông, ODF, kẹp…)', 'cái', function (f) { return /mang xong|closure|\bodf\b|hop noi|splice|pigtail|kep cap quang|ftth box|termination (box|unit)|\bfdt\b|\bfat\b|splitter|bang bao hieu|warning tape|notice optical|fiber termination/.test(f); }, 'ngv'],
    ['shel', 'Nhà trạm, container, shelter', 'nhà', function (f) { return /nha tram|nha container|container house|\bcontainer\b|minishelter|\bshelter\b|nha may no|house for generator|nha lap ghep|prefab/.test(f) && !accH(f); }, 'ht'],
    ['tower', 'Cột anten / tower', 'cột', function (f) { return /cot anten|cot an ten|cot ang ten|cot tu dung|cot day co|monopole|\btower\b|guyed mast|self.?support(ing)? (tower|mast)|cot thep goc/.test(f) && !/bulong|bolt|ban|plate|moc|hook|op |veneer|kep|clamp|foundation|parts|day co|cap|cable|thang|ladder|den|light|pin|tiep dia/.test(f); }, 'ht'],
    ['twp', 'Đốt cột, phụ kiện cột anten (bản ốp, kim thu sét, bulông móng, cáp thép…)', 'cái', function (f) { return /dot cot|typical (piece|segment)|kim thu set|lightning rod|ban op|ban dinh vi|anchor bolt template|mat bich|flange head|bu ?long mong|foundation bolt|cleat hook|moc neo|\bday co\b|guy wire|tang do|cap thep|steel cable|khoa cap|cable lock|thang leo|ladders? for mast|chong xoan|rotation security|top cover plate|vong op|ga chong xoan/.test(f); }, 'ht'],
    ['cow', 'Trạm phát sóng lưu động (xe, COW)', 'xe', function (f) { return /xe phat song|car tower|cell on wheel|\bcow\b|mobile bts/.test(f); }, 'ran'],
    ['txc', 'Card, khung truyền dẫn (OSN, WDM, IP, OLT…)', 'card', function (f) { return /line ?card|juniper|\bmpc\d|\bmx ?\d{3,4}\b|\btn\d\d|\btn[a-z]\d|\btme\d|\btng\d|\btns\d|osn ?\d|optix|\bwdm|dwdm|line service|multiplexing board|tributary|cfpt|\bcfp\b|ne ?40|ne ?\d\de|\batn ?\d|\bcx ?600|xgpon|\bgpon|zxa10|\bc6[0-9]0\b|subrack|chassis|\bccf0|communication control module/.test(f); }, 'tx'],
    ['card', 'Card / module BBU', 'card', function (f) { return /\bcard\b|\bboard\b|\bubbp|\bumpt|\blmpt|\bwmpt|\bgtmu|\bupeu|\bueiu|\busch|\bubri|\blbbp|\bwbbp|\bgbbp|\butrp|\bumdu|\bbpn\d?\b|\bbpq|\bvbp|\bppc\d*|\bfs\d\b|\bcc\d+\b|fan unit|fan module|khoi quan ly quat|\bubfa|\bfan[a-z]?_|bo nguon bang tan|bang mach|bo mach|main processing|baseband processing|umpt|ubbp|lmpt|wmpt|fan array|\bfa\dc\b|\bfs\d[a-z]\b|\bsa\d\b|\bpm\d\b|module canh bao|module chuyen mach|modul nguon/.test(f) && !/(^|[^a-z])(cap|cable|day)([^a-z]|$)/.test(f); }, 'ran'],
    ['bbu', 'BBU / khối băng gốc', 'bộ', function (f) { return /khoi xu ly bang goc|baseband unit|khung chua bbu|plug-in box|\bbbu ?\d{3,4}\b|\bbbu box|\bbbu\b/.test(f) && !accH(f); }, 'ran'],
    ['mw', 'Viba (IDU/ODU/anten viba)', 'bộ', function (f) { return /viba|microwave|\bidu\b|\bodu\b|\bnec\b|pasolink|\brtn ?\d|ipasolink|\bmw\b|\brfu\b|ceragon|all indoor|\bip ?20|\bip ?10\b/.test(f) && !/(^|[^a-z])(cap|cable|day|kep|ga)([^a-z]|$)/.test(f); }, 'tx'],
    ['rru', 'RRU / AAU (khối thu phát)', 'khối', function (f) { return /khoi thu phat|radio unit|active antenna unit|rru|\baau ?\d*|[^a-z]a9\d{3}|\bmrfu|\bdrfu|\bwrfu|\blrfu|\bgrfu|\brrh\b|\br88\d\d|\br92\d\d/.test(f) && !accH(f) && !/^(sfp|xfp|module quang|optical)|\bodu\b|\bidu\b|\bnec\b/.test(f); }, 'ran'],
    ['tx0', 'Module quang (SFP/XFP)', 'module', function (f) { return /\bsfp|\bxfp|\bqsfp|optical transceiver|bidi transceiver|module quang|modul quang|optical module/.test(f) && !/khoi thu phat|\brru|radio unit|\br88\d\d|\br92\d\d|zxsdr r/.test(f); }, 'tx'],
    ['ant', 'Antenna trạm', 'cái', function (f) { return /anten|antenna/.test(f) && !/gps|gnss|(^|[^a-z])(cap|cable|parts?|bulong|bolt|plate|jumper|kep|clamp|kit|bracket)([^a-z]|$)|ret control|rcu|aisg|dieu khien|day nhay|ban dinh vi|ban op|ga lap|mounting|thanh chong|strut|cot (tu dung|day co)|foundation|self supporting/.test(f); }, 'ran'],
    ['racc', 'Phụ kiện vô tuyến (cáp RRU, gá, kẹp, RET, GPS, diplexer…)', 'cái', function (f) { return /rru|bbu|anten|antenna|aau|\bret\b|aisg|gps|gnss|\bbts\b|enodeb|nodeb|diplexer|triplexer|combiner|duplexer|bo ghep kenh|\btma\b|\bodm ?\d|hop phan phoi nguon|\bdpu\d|\bdcpd/.test(f); }, 'ran'],
    ['tx', 'Truyền dẫn quang / IP (OLT, switch, router…)', 'bộ', function (f) { return /\bolt\b|\bsfp|transceiver|optical module|module quang|\bswitch\b|\brouter\b|\bptn\b|\bdwdm|\bsdh\b|\bmux\b|\botn\b|\bipran|\batn ?\d|\bcx ?600|\bnetengine|\bmetro\b/.test(f); }, 'tx'],
    ['cab', 'Tủ thiết bị / tủ nguồn', 'tủ', function (f) { return /cabinet|\btu (nguon|thiet bi|outdoor|indoor|ac|dc|bts|rack|dien)|\bapm ?\d|\btp48|\brack\b|shelter/.test(f) && !accH(f); }, 'pw'],
    ['bat', 'Ắc quy', 'bình/khối', function (f) { return /battery|acquy|ac quy|\bbin\b.*lithium|lithium/.test(f) && !/(^|[^a-z])(cap|cable|day|gia|rack|ke|sac|charger)([^a-z]|$)/.test(f); }, 'pw'],
    ['gsp', 'Phụ tùng máy phát (lọc dầu, lọc gió…)', 'cái', function (f) { return /loc (dau|nhien lieu|gio|nhot)|filter|phu tung|spare part|bugi|day curoa|belt|stator|rotor|bom cao ap|injection|piston|xi lanh|cylinder|gioang may/.test(f); }, 'pw'],
    ['pwr', 'Nguồn: rectifier, máy phát, ATS, solar, điều hoà', 'bộ', function (f) { return /rectifier|\brect\b|power supply|\bpsu\b|inverter|may phat|generator|\bats\b|solar|pin mat troi|\bups\b|module nguon|bo nguon|dieu hoa|air condition|\bmppt|\bavr\b|on ap|\bmpd\b|genset|\bkva\b|power unit|dc power|khoi nguon|phan phoi nguon|distribution (unit|module)|\bzxdt/.test(f) && !ACC.test(f) && !/stator|bom cao ap|injection|loc |filter/.test(f); }, 'pw'],
    ['elec', 'Vật tư điện (aptomat, CB, công tơ, cầu chì, ổ cắm, chống sét nguồn…)', 'cái', function (f) { return /\bcb\b|cong to|energy meter|\bkwh\b|aptomat|attomat|circuit breaker|\bmcb\b|\bmccb\b|cau chi|\bfuse\b|o cam|socket|cau dao|contactor|cong tac|surge protect|cat loc set|chong set nguon|\bmov\b|bon dau|oil tank|box oil|tu dien|electrical (box|panel)/.test(f); }, 'pw'],
    ['civ', 'Vật tư xây dựng trạm, tiếp địa (thép, hàng rào, thuốc hàn, bảng đồng…)', 'cái', function (f) { return /hang rao|fence|\bthep\b|steel|tiep dia|grounding|earthing|thuoc han|welding drug|corrodent|bang lo cap|cable entry|wiring window|thang cap|cable tray|cable ladder|mang cap|xi mang|cement|gach|cat vang|\bson\b|paint|coc (dat|thep|dong|tiep dia)|bang dong|busbar|copper (board|bar|panel)|bien ten|bien bao tram/.test(f); }, 'ht'],
    ['tool', 'Công cụ, dụng cụ, máy đo', 'cái', function (f) { return /may do|otdr|theodolite|kinh vi|dung cu|\btools?\b|(power|optical|test|multi|clamp|insulation|earth|resistance|signal|light|laser|volt|ohm)[ -]?meter|\btester\b|fusion splicer|may han|spectrum analy/.test(f) && !/energy meter|cong to|\bkwh\b/.test(f); }, 'cc'],
    ['feed', 'Feeder, jumper RF, connector RF', 'm', function (f) { return /feeder|coaxial|\bcoax\b|rf (cable|jumper)|jumper|\b1 2 ?(in\b|")|\b7 8\b|\b7 16\b|\bn ?type\b|\bdin\b|\brg ?\d|\bqma\d/.test(f) && !/quang|optical|fib(er|re)|patch ?cord|\b(lc|sc|fc|st)\b/.test(f); }, 'ran'],
    ['jmpq', 'Dây nhảy quang (patch cord)', 'sợi', function (f) { return /day nhay quang|patch ?cord|optical (patch|jumper|cable assembl)|fiber jumper|day quang|soi quang|single mode fiber \(|\b(lc|sc|fc|st|e2000)( (pc|upc|apc))? ?[- ] ?(lc|sc|fc|st|e2000)\b/.test(f); }, 'tx'],
    ['pcab', 'Cáp nguồn, dây điện', 'm', function (f) { return /cap nguon|power (cable|cord|wire)|day (dien|nguon|ac|dc|dat)\b|cap dien|electric (wire|cable)|mm2|mm\^2|cu xlpe|\bvav\b|\bcvv\b|\bcxv\b|cap nhom|aluminium cable|aluminum cable|cap dong (tran|boc)|day dong boc|ground(ing)? (wire|cable)|\bwire\b/.test(f); }, 'pw'],
    ['ccab', 'Cáp đồng, cáp mạng, cáp luồng', 'm', function (f) { return /cap dong|copper cable|\butp\b|\bftp\b|\bstp\b|cat ?5|cat ?6|cap mang|network cable|lan cable|cap internet|cap luong|\be1\b|trunk cable|signal cable|cable.?assy|rs ?232|serial|twisted.?pair/.test(f); }, 'tx'],
    ['cable', 'Cáp, dây khác', 'm', function (f) { return /(^|[^a-z])(cap|cable|day|wire|patch)([^a-z]|$)/.test(f); }, 'kv'],
    ['core', 'Core / CNTT (máy chủ, lưu trữ…)', 'bộ', function (f) { return /\bserver\b|storage|luu tru|\be9000|core network|\bmsc\b|\bmme\b|\bhlr\b|\bsgsn|\bggsn|hard ?disk|\bhdd\b|\bssd\b|may chu|\bmemory\b|\bram\b|\bcpu\b/.test(f); }, 'it'],
    ['pc', 'Máy tính, thiết bị văn phòng', 'cái', function (f) { return /may tinh|laptop|desktop|notebook|\bimac\b|\bpc\b|man hinh|monitor|may in|printer|scanner|may scan|may chieu|projector|tivi|smart tv|\bban (lam viec|hop|giao dich)|\bghe\b|chair|sofa|tu (sat|ho so)|ket sat|cay nuoc|webcam|may fax|may photo/.test(f); }, 'vp'],
    ['veh', 'Xe, phương tiện vận tải', 'xe', function (f) { return /\bxe (may|tai|o to|ban tai|nang)|motorbike|motobike|motorcycle|truck|pickup|\bcar\b|vehicle|lop o to|\btyre|\btire\b|nissan|toyota|hyundai|honda wave/.test(f); }, 'pt'],
    ['term', 'Đầu cuối, SIM, CPE, ONT', 'cái', function (f) { return /\bsim\b|sim card|\bcpe\b|modem|handset|dien thoai|\bftth\b|set.?top|\bstb\b|scratch card|the cao|\bonu\b|\bont\b|\bmifi|router wifi|wifi router/.test(f); }, 'kd'],
    ['acc', 'Phụ kiện, vật tư phụ', 'cái', function (f) { return ACC.test(f); }, 'kv']];
  var TECH = [['5G', /\b5g\b|\bnr\b|gnodeb|v9200|aau/], ['4G', /\b4g\b|\blte\b|\bfdd\b|\btdd\b|enodeb|\blmpt|\blbbp|\blrfu|bbu5900/], ['3G', /\b3g\b|\bumts|\bwcdma|\bnodeb|\bwbbp|\bwrfu|\bwmpt|\bgu ?360|\bu ?2100/], ['2G', /\b2g\b|\bgsm|\bgtmu|\bdrfu|\bgbbp|\bgrfu|bts ?3012|\bgu ?360|\bs9000|\bb8 ?900/]];
  var VEND = [['Huawei', /huawei|\brru ?3\d{3}|\bbbu ?3\d{3}|\bbbu ?3910|\bbbu ?5900|\brru ?39\d\d|\bdbs ?3900|\bubbp|\bumpt|\bupeu|\bubfa|\brtn/], ['ZTE', /(^|[^a-z0-9])zte([^a-z0-9]|$)|\brru ?r?8\d{3}|\brru_?r9\d{3}|[^a-z]a9\d{3}|zxsdr|zxran|\br88\d\d|\br92\d\d|\bv9200|\bb8200|\bbpn|\bbpq|\bvbp/], ['Nokia', /nokia|flexi|airscale/], ['Ericsson', /ericsson|\brbs ?\d/], ['NEC', /(^|[^a-z])nec([^a-z]|$)|pasolink/]];
  function domOfGroup(g) { var f = fold(g); if (!f) return 'kv'; return /vo tuyen|\bbts\b|\bran\b|radio/.test(f) ? 'ran' : /truyen dan|viba|microwave/.test(f) ? 'tx' : /cap quang|ngoai vi/.test(f) ? 'ngv' : /co dien|nguon/.test(f) ? 'pw' : /ha tang|\bcot\b/.test(f) ? 'ht' : /tong tram|tong dai|\bcore\b/.test(f) ? 'tt' : /cntt|\bit\b|tin hoc/.test(f) ? 'it' : /^ccdc|cong cu|dung cu/.test(f) ? 'cc' : /ttbvp|van phong/.test(f) ? 'vp' : /ptvt|phuong tien|van tai/.test(f) ? 'pt' : /dau cuoi|sim|kinh doanh/.test(f) ? 'kd' : 'kv'; }
  var NEUTRAL = /^(bat|pcab|cab|tx0|jmpq|ccab|feed|elec|acc|cable|tool|pc|civ|polec)$/;
  var XMAP = { 'card@tx': 'txc', 'txc@ran': 'card', 'card@tt': 'ttc', 'txc@tt': 'ttc', 'tx@tt': 'ttc', 'bbu@tt': 'ttc', 'core@tt': 'ttc', 'card@it': 'core', 'txc@it': 'core', 'tx@it': 'core', 'core@vp': 'pc', 'twp@ngv': 'ngacc', 'shel@pw': 'cab', 'tower@ngv': 'polec' };
  
const CC={};NT.fold=fold;NT.cls=(name,grp)=>{const key=name+'\u0001'+grp;if(CC[key])return CC[key];const f=fold(name).replace(/[_\/]+/g,' ');let c=null;for(let i=0;i<CATS.length;i++)if(CATS[i][3](f)){c=CATS[i];break}
 let vd='';for(const x of VEND)if(x[1].test(f)){vd=x[0];break}const dom=domOfGroup(grp);let k=c?c[0]:'oth';if(c&&dom!=='kv'&&c[4]!==dom&&c[4]!=='kv'&&!NEUTRAL.test(c[0]))k=XMAP[c[0]+'@'+dom]||'gx';return CC[key]={k,vd}};
NT.pl=(R,cfg)=>{const rate=+cfg.rate||1,Y=new Date(cfg.cut+'T00:00:00').getFullYear(),U=v=>(+v||0)/rate,tabs=[],add=(n,t,a)=>tabs.push({n,t,a}),mi=R.filter(r=>r.V===NV.mi);
 add('TH','Tổng hợp tồn (USD): năm phân loại | loại hàng × nhóm VTTB',A('Năm | Loại hàng',P(R,r=>r.X+' | '+r.V,r=>r.W,r=>U(r.L))));
 add('PL04','PL04 Thời điểm nhập kho (USD, theo số tiền bàn giao)',A('Loại hàng',P(R,r=>r.V,r=>r.ad,r=>U(r.M))));
 add('PL7','PL7 Tồn hàng mới theo nhóm × thời điểm nhập (USD)',A('Nhóm VTTB',P(mi,r=>r.W,r=>r.ad,r=>U(r.M))));
 const hd=P(R.filter(r=>r.u),r=>r.u,r=>r.V,r=>U(r.M));hd.rk.sort((a,b)=>Object.values(hd.m.get(b)).reduce((x,y)=>x+y,0)-Object.values(hd.m.get(a)).reduce((x,y)=>x+y,0));
 add('PL_HopDong','PL hợp đồng: số hợp đồng không trùng × loại hàng (USD)',A('Số hợp đồng',hd));
 [Y-1,Y].forEach(y=>add('PL_HD_'+y,'Hợp đồng nhập năm '+y+': hàng mới, trong/quá hạn (USD)',A('Số hợp đồng',P(mi.filter(r=>r.AB===y&&r.X===String(y)),r=>r.u||'(trống)',r=>r.ac,r=>U(r.M)))));
 return tabs};
NT.enr=rs=>rs.forEach(r=>{if(r.ck===undefined){const c=NT.cls(r.nm,r.W);r.ck=c.k;r.vd=c.vd;r.qn=r.V===NV.mi?r.q:0;r.qu=r.V===NV.noc?r.q:0}});
NT.dev=(rs,rate)=>{const k=+rate||1,Q=(cat,vd,f)=>{let s=0;rs.forEach(r=>{if(r.ck===cat&&(vd==null||r.vd===vd))s+=f==='v'?r.L/k:f==='qn'?r.qn:f==='qu'?r.qu:r.q});return s},up=(c,vd)=>{const q=Q(c,vd);return q?Q(c,vd,'v')/q:0};
 const vends=['Huawei','ZTE','Nokia','Ericsson'].filter(v=>Q('bbu',v)+Q('rru',v)>0),antN=Q('ant',null,'qn'),antA=Q('ant')-Q('ant',null,'qu'),upA=up('ant');
 const V=vends.map(vd=>{const b=Q('bbu',vd),r=Q('rru',vd),c=Q('card',vd),bn=Q('bbu',vd,'qn'),rn=Q('rru',vd,'qn'),bu=b-Q('bbu',vd,'qu'),ru=r-Q('rru',vd,'qu'),sN=Math.floor(Math.min(bn,rn/3)),sA=Math.floor(Math.min(bu,ru/3)),ub=up('bbu',vd),ur=up('rru',vd);
  return{vd,b,r,c,bn,rn,sN,sA,vB:Q('bbu',vd,'v'),vR:Q('rru',vd,'v'),vC:Q('card',vd,'v'),vSet:sA*(ub+3*ur+3*upA),surB:Math.max(0,bu-sA),surR:Math.max(0,ru-3*sA),vSur:Math.max(0,bu-sA)*ub+Math.max(0,ru-3*sA)*ur,lim:sA===Math.floor(bu)?'BBU':'RRU'}});
 const b2=rs.filter(r=>r.ck==='bat'&&/(^|[^0-9.])2 ?v([^a-z0-9]|$)/i.test(fold(r.nm))).reduce((a,r)=>a+r.q,0);
 return{V,oth:{b:Q('bbu',''),r:Q('rru','')},antN,antA,sumA:V.reduce((a,x)=>a+x.sA,0),sumN:V.reduce((a,x)=>a+x.sN,0),Q,b2}};
NT.calc=(S,cfg)=>{
 const rate=+cfg.rate||1,cut=new Date(cfg.cut+'T00:00:00'),Y=cut.getFullYear(),since=(cut-new Date(Y,0,1))/DAY,U=v=>(+v||0)/rate;
 const LK={...(S.cat||{lk:{}}).lk,...cfg.ov.lk},NH={...(S.cat||{nh:{}}).nh,...cfg.ov.nh},tabs=[],bad={};
 const ix=s=>n=>col((S[s]||{m:{}}).m,n),rowsOf=s=>(S[s]||{rows:[]}).rows;
 const SH={...((S.mem||{}).lo||{})},okd=d=>d&&d.getFullYear()>=2000&&d.getFullYear()<=Y+1?d:null;
 /* Tiến độ lô hàng: ngày nhập kho lô = cột AA (chỉ số 26, "ngày nhập kho TCT") – đúng như tool mẫu VLOOKUP cột AA.
    Tiêu đề "Finish date" lặp ở N, P, T, V, AA, AD, AG nên KHÔNG tra theo tên. Các cột khác chỉ để gợi ý khi AA trống. */
 const TDC=[26,32,29,21,15,19,13,11,5],TDL={};
 {const g=ix('td'),a=g('shipment code');
  rowsOf('td').forEach(r=>{const k=String(r[a]).trim();if(!k)return;const d=okd(dt(r[26]));if(d)SH[k]=d.getTime();else{for(const j of TDC.slice(1)){const e=okd(dt(r[j]));if(e){TDL[k]={t:e.getTime(),c:j};break}}}})}
 Object.entries(cfg.ov.lo||{}).forEach(([k,v])=>{if(v)SH[k]=v});
 /* ngày nhập theo HĐ từ Import_Summary NCC: (HĐ + mã hàng) sớm nhất, nếu không có thì HĐ sớm nhất */
 const IC={...((S.mem||{}).ic||{})},IP={},INF={},LI=new Set();
 {const g=ix('ncc'),a=g('date import'),u=g('contract no'),e=g('equipment code'),n=g('equipment name'),sp=g('supplier name'),pr=g('price'),f=g('contract date');
  rowsOf('ncc').forEach(r=>{const c=String(r[u]).trim(),d=okd(dt(r[a]));if(!c||c[0]==='('||!d)return;const t=d.getTime(),cd=String(r[e]).trim(),k=c+'|'+cd;if(!(k in IP)||t<IP[k])IP[k]=t;if(!(c in IC)||t<IC[c])IC[c]=t;
   const o=INF[cd];if(cd&&(!o||t>o.t))INF[cd]={t,nm:String(r[n]).split('|')[0].trim(),sp:String(r[sp]).trim(),c,p:+String(r[pr]).replace(/,/g,'')||0,cdt:(x=>x?p2(x.getDate())+'/'+p2(x.getMonth()+1)+'/'+x.getFullYear():'')(dt(r[f]))}})}
 const PN=new Set((S.mem||{}).pn||[]),SN=new Set((S.mem||{}).sn||[]);{let g=ix('pnk'),a=g('import code');rowsOf('pnk').forEach(r=>PN.add(String(r[a]).trim()));g=ix('sn');a=g('serial number');rowsOf('sn').forEach(r=>SN.add(sk(r[a])))}
 const IL={},IN={},ne={n:0,v:0},bd=(t,k,v,e,x2)=>{const x=bad[t+k]=bad[t+k]||{t,k,n:0,v:0,e,...(x2||{})};x.n++;x.v+=v};
 const g=ix('ct'),c={k:g('code warehouse'),cl:g('code list'),sn:g('serial'),st:g('status'),l:g('amount at cost'),m:g('to hand over'),o:g('storage period'),p:g('business process'),t:g('shipment no'),u:g('contract no'),nm:g('name list'),q:g('number')},R=[];
 rowsOf('ct').forEach(r=>{const kho=String(r[c.k]).trim();if(!kho||kho[0]==='(')return;
  const st=String(r[c.st]).trim(),cd=String(r[c.cl]).trim(),T=String(r[c.t]).trim(),U1=String(r[c.u]).trim(),ser=sk(r[c.sn]),L=+r[c.l]||0;
  const V=LK[kho+st]||NV.un,h=NH[cd]||{},W=h.g||NV.un,Z=h.t&&h.t!=='0'?h.t:NV.un,d=SH[T]||IP[U1+'|'+cd]||IC[U1],AB=d?new Date(d).getFullYear():0,O=+r[c.o]||0,mi=V===NV.mi||V===NV.noc;
  const sx=kho.split('_').slice(1).join('_')+'|'+st,w2=String(r[c.nm]).toLowerCase().split(/\s+/).slice(0,2).join(' ');if(V===NV.un)bd('lk',kho+st,L,sx,{kho,st});else{const o=IL[sx]=IL[sx]||{};o[V]=(o[V]||0)+1}if(W===NV.un||Z===NV.un)bd('nh',cd,L,w2,{nm:String(r[c.nm]).trim()});else{const o=IN[w2]=IN[w2]||{};o[W+'|'+Z]=(o[W+'|'+Z]||0)+1}if(!d)bd('lo',T||'(trống)',L,U1);else if(!SH[T])LI.add(T);
  const X=V===NV.un?'Kiểm tra':([NV.th,NV.kcs,NV.tl].includes(V)&&(AB===Y||O<=since)&&(PN.has(String(r[c.p]).trim())||(ser&&SN.has(ser)))?'Thu hồi':AB===Y?String(Y):String(Y-1)),age=mi&&d?(cut-d)/DAY:O,b=kho.split('_')[0].toUpperCase();
  if([NV.th,NV.kcs,NV.tl].includes(V)&&O<=since&&X!=='Thu hồi'){ne.n++;ne.v+=L}R.push({cd,nm:String(r[c.nm]).trim(),q:+r[c.q]||0,kho,br:BR.includes(b)?b:'KHO CÔNG TY',V,W,Z,X,AB,L,M:+r[c.m]||0,u:String(r[c.u]).trim(),
   ad:!AB?'Chưa rõ':AB>=Y?String(Y):AB>=Y-1?String(Y-1):AB>=Y-2?String(Y-2):'Trước '+(Y-2),ac:!d?'Chưa rõ':(cut-d)/DAY<181?'Trong hạn':'Quá hạn',
   ae:age<91?'Dưới 3 tháng':age<182?'3-6 tháng':age<365?'6-12 tháng':age<1095?'1-3 năm':'Trên 3 năm'})});
 const add=(n,t,a)=>tabs.push({n,t,a}),mi=R.filter(r=>r.V===NV.mi);
 add('Ton_ChiNhanh','Tồn kho (USD) theo loại hàng × chi nhánh/kho công ty',A('Loại hàng',P(R,r=>r.V,r=>r.br,r=>U(r.L))));
 add('TH','Tổng hợp tồn (USD): năm phân loại | loại hàng × nhóm VTTB',A('Năm | Loại hàng',P(R,r=>r.X+' | '+r.V,r=>r.W,r=>U(r.L))));
 add('PL04','PL04 Thời điểm nhập kho (USD, theo số tiền bàn giao)',A('Loại hàng',P(R,r=>r.V,r=>r.ad,r=>U(r.M))));
 add('PL7','PL7 Tồn hàng mới theo nhóm × thời điểm nhập (USD)',A('Nhóm VTTB',P(mi,r=>r.W,r=>r.ad,r=>U(r.M))));
 const hd=P(R.filter(r=>r.u),r=>r.u,r=>r.V,r=>U(r.M));hd.rk.sort((a,b)=>Object.values(hd.m.get(b)).reduce((x,y)=>x+y,0)-Object.values(hd.m.get(a)).reduce((x,y)=>x+y,0));
 add('PL_HopDong','PL hợp đồng: số hợp đồng không trùng × loại hàng (USD)',A('Số hợp đồng',hd));
 [Y-1,Y].forEach(y=>add('PL_HD_'+y,'Hợp đồng nhập năm '+y+': hàng mới, trong/quá hạn (USD)',A('Số hợp đồng',P(mi.filter(r=>r.AB===y&&r.X===String(y)),r=>r.u||'(trống)',r=>r.ac,r=>U(r.M)))));
 /* đối soát tồn tiền */
 const bk={};R.forEach(r=>bk[r.kho]=(bk[r.kho]||0)+r.L);const tk={};{const g=ix('tt'),a=g('warehouse code'),b=g('amount');rowsOf('tt').forEach(r=>{const k=String(r[a]).trim();if(k&&typeof r[b]==='number')tk[k]=(tk[k]||0)+r[b]})}
 const n0=v=>Math.round(v).toLocaleString('en-US'),DX={};
 {const g=ix('tt'),a=g('warehouse code'),b=g('amount'),cc=g('code'),qq=g('quantity'),pn=g('product name'),C=new Map();
  const gt=(k,c,nm)=>{const key=k+'\u0001'+c;let o=C.get(key);if(!o)C.set(key,o={k,c,nm:'',cq:0,cv:0,tq:0,tv:0});if(nm&&!o.nm)o.nm=nm;return o};
  R.forEach(r=>{const o=gt(r.kho,r.cd,r.nm);o.cq+=r.q;o.cv+=U(r.L)});
  rowsOf('tt').forEach(r=>{const k=String(r[a]).trim();if(!k||typeof r[b]!=='number')return;const o=gt(k,String(r[cc]).trim(),String(r[pn]).trim());o.tq+=+r[qq]||0;o.tv+=U(r[b])});
  const byC=new Map();C.forEach(o=>{o.d=o.cv-o.tv;let t=byC.get(o.c);if(!t)byC.set(o.c,t={p:0,n:0});if(o.d>0)t.p+=o.d;else t.n-=o.d});
  C.forEach(o=>{const t=byC.get(o.c),m=Math.min(t.p,t.n);o.tr=m<1?0:(o.d>0?o.d*m/t.p:o.d*m/t.n);o.rest=o.d-o.tr;
   o.cat=Math.abs(o.d)<1?'':Math.abs(o.tr)>=Math.abs(o.d)-1?'chuyển/xếp khác kho':o.tq===0&&o.tv===0?'mã chỉ có ở tồn biểu đồ':o.cq===0&&o.cv===0?'mã chỉ có ở tồn tiền':Math.abs(o.cq-o.tq)>0.5?'lệch số lượng':'lệch đơn giá/thành tiền';
   const k=o.k,x=DX[k]=DX[k]||{p:{},n:0,top:[]};
   if(Math.abs(o.tr)>=1)x.p['chuyển/xếp khác kho']=(x.p['chuyển/xếp khác kho']||0)+o.tr;
   if(Math.abs(o.rest)>=1){const c=o.tq===0&&o.tv===0?'mã chỉ có ở tồn biểu đồ':o.cq===0&&o.cv===0?'mã chỉ có ở tồn tiền':Math.abs(o.cq-o.tq)>0.5?'lệch số lượng':'lệch đơn giá/thành tiền';x.p[c]=(x.p[c]||0)+o.rest}
   if(Math.abs(o.d)>=1){x.n++;x.top.push(o)}});
  DX.C=[...C.values()].filter(o=>Math.abs(o.d)>=1).sort((x,y)=>Math.abs(y.d)-Math.abs(x.d));
 }
 const MS=(S.mem||{}).ds||{},cm=mk(cut),pm=Object.keys(MS).filter(k=>k<cm).sort().pop(),PV=pm?MS[pm]:null,DM={};
 const ACT={'chuyển/xếp khác kho':'Đối chiếu phiếu chuyển kho; hai báo cáo xếp cùng mã vào kho khác nhau, cần điều chỉnh về một kho','mã chỉ có ở tồn biểu đồ':'Kiểm tra hàng nhập hoặc phát sinh sau ngày chốt tồn tiền, hoặc chưa hạch toán tài chính','mã chỉ có ở tồn tiền':'Kiểm tra hàng đã xuất, thanh lý còn nằm trên sổ tiền hoặc chưa cập nhật WMS','lệch số lượng':'Kiểm kê số lượng thực tế các mã đứng đầu','lệch đơn giá/thành tiền':'Kiểm tra đơn giá và tỷ giá giữa hai báo cáo'};
 const ds=[...new Set([...Object.keys(bk),...Object.keys(tk)])].sort().map(k=>{const x=U(bk[k]),y=U(tk[k]),d=x-y,w=!(k in bk)?'Chỉ có ở tồn tiền':!(k in tk)?'Chỉ có ở tồn biểu đồ':Math.abs(d)>100&&Math.abs(d)>0.005*Math.max(Math.abs(x),Math.abs(y))?'Lệch lớn':'Khớp';DM[k]=d;
  if(w==='Khớp')return[k,x,y,d,w,'','','',0,'','',''];
  const X=DX[k]||{p:{},n:0,top:[]},ps=Object.entries(X.p).sort((a,b)=>Math.abs(b[1])-Math.abs(a[1])),tr=X.p['chuyển/xếp khác kho']||0,rem=d-tr;
  const cause=w==='Chỉ có ở tồn tiền'?'Mã kho không có ở tồn biểu đồ (kho đã đóng, đổi tên hoặc chưa mapping)':w==='Chỉ có ở tồn biểu đồ'?'Mã kho không có ở tồn tiền (kho mới, đổi tên hoặc chưa mapping)':ps.length?ps[0][0]+' ('+n0(ps[0][1])+' USD)':'Chênh nhỏ, làm tròn';
  const top=X.top.sort((a,b)=>Math.abs(b.d)-Math.abs(a.d)).slice(0,3).map(o=>o.c+(o.nm?' '+o.nm.slice(0,28):'')+': '+n0(o.d)).join('; ');
  const pv=PV&&k in PV?PV[k]:null,tn=pv===null?'Chưa có tháng trước để so sánh':'Tháng '+pm+': '+n0(pv)+(Math.abs(d)>Math.abs(pv)*1.05?' → tăng thêm':Math.abs(d)<Math.abs(pv)*0.95?' → giảm':' → gần như không đổi');
  const dub=pv===null?'Nếu không xử lý, tháng sau còn lệch khoảng '+n0(d)+' USD (giữ nguyên hiện tại)':'Nếu không xử lý, tháng sau lệch khoảng '+n0(d+(d-pv))+' USD (theo xu hướng); xử lý đúng nguyên nhân chính còn khoảng '+n0(rem)+' USD';
  return[k,x,y,d,w,X.n,cause,ps.map(p=>p[0]+' '+n0(p[1])).join('; '),rem,top,tn+'. '+dub,ps.length?ACT[ps[0][0]]:'Đối chiếu chi tiết từng mã']});
 {const mm0={...MS};mm0[cm]=DM;const ks=Object.keys(mm0).sort().slice(-6);const o={};ks.forEach(k=>o[k]=mm0[k]);S.mem=S.mem||{};S.__ds=o}
 const dsw=ds.filter(r=>r[4]!=='Khớp'),gross=dsw.reduce((a,r)=>a+Math.abs(r[3]),0),net=dsw.reduce((a,r)=>a+r[3],0),remG=dsw.reduce((a,r)=>a+Math.abs(r[8]),0);
 add('DoiSoat','Đối soát tồn biểu đồ và tồn tiền theo mã kho (USD)',[['Mã kho','Tồn biểu đồ','Tồn tiền','Chênh lệch','Cảnh báo','Số mã lệch','Nguyên nhân chính','Phân tích nguyên nhân (USD)','Lệch còn lại sau khi loại chuyển kho','Top mã gây lệch (USD)','So sánh tháng trước và dự báo','Việc cần làm'],...ds,['Tổng',...[1,2,3].map(i=>ds.reduce((a,r)=>a+r[i],0)),'','','','','','','','','']]);
 add('DoiSoat_Ma','Đối soát chi tiết theo kho × mã hàng (USD), chỉ các dòng lệch',[['Mã kho','Mã hàng','Tên hàng','SL biểu đồ','SL tồn tiền','USD biểu đồ','USD tồn tiền','Chênh lệch','Phân loại'],...DX.C.slice(0,5000).map(o=>[o.k,o.c,o.nm,o.cq,o.tq,o.cv,o.tv,o.d,o.cat])]);
 /* theo tháng */
 const mon={},am=(m,k,v)=>{(mon[m]=mon[m]||{x:0,tl:0,th:0})[k]+=v};
 {let g=ix('ex');rowsOf('ex').forEach(r=>{const d=dt(r[g('date out')]),re=String(r[g('reason')]);if(!d||/chuyển kho|export warehouse from/i.test(re)||/lost/i.test(re)||(cfg.con&&!String(r[g('construction code')]).trim()))return;am(mk(d),'x',U(r[g('cash total')]))});
  g=ix('tl');rowsOf('tl').forEach(r=>{const d=dt(r[g('date export')]);if(d)am(mk(d),'tl',U(r[g('cash total')]))});
  g=ix('pnk');rowsOf('pnk').forEach(r=>{const d=dt(r[g('date import')]);if(d)am(mk(d),'th',U(r[g('cash total')]))})}
 const mm={...((S.mem||{}).mon||{}),...mon};add('Thang','Theo tháng (USD): xuất ra trạm (bỏ chuyển kho), thanh lý, nhập thu hồi',[['Tháng','Xuất ra trạm','Thanh lý','Nhập thu hồi'],...Object.keys(mm).sort().map(k=>[k,mm[k].x,mm[k].tl,mm[k].th])]);
 /* PL03 nhập NCC từ trước đến nay */
 const yr={},CV={},CQ={},CQq={},NM={};let inv=0,invY=0;{const g=ix('ncc');rowsOf('ncc').forEach(r=>{const d=dt(r[g('date import')]),v=U(r[g('cash total')]);if(!d)return;inv+=v;NM[mk(d)]=(NM[mk(d)]||0)+v;{const hc=String(r[g('contract no')]).trim(),cd=String(r[g('equipment code')]).trim();if(hc&&hc[0]!=='(')CV[hc]=(CV[hc]||0)+v;if(cd){CQ[cd]=(CQ[cd]||0)+v;CQq[cd]=(CQq[cd]||0)+(+r[g('quantity real')]||0)}};if(d.getFullYear()===Y)invY+=v;const y=d.getFullYear(),o=yr[y]=yr[y]||[0,0];o[0]++;o[1]+=v})}
 add('PL03','PL03 Tổng nhập hàng hóa từ NCC từ trước đến nay (USD)',[['Năm nhập','Số dòng','Giá trị'],...Object.keys(yr).sort().map(y=>[y,yr[y][0],yr[y][1]]),['Tổng',Object.values(yr).reduce((a,o)=>a+o[0],0),inv]]);
 const sum=f=>R.filter(f).reduce((a,r)=>a+U(r.L),0),tot=sum(()=>true),q=(a,b)=>b?a/b*100:0,sy=String(Y);
 const kpi=[['Tổng tồn kho (USD)',tot,'',0],['KPI01 Tồn/Đầu tư',q(tot,inv),'%',+cfg.k1],['KPI02 Tồn (trừ UCTT, không PS năm nay)/Đầu tư',q(R.filter(r=>r.V!==NV.noc&&r.AB!==Y).reduce((a,r)=>a+U(r.L),0),inv),'%',+cfg.k2],
  ['KPI03 Mới quá hạn/Nhập mới năm '+Y,q(R.filter(r=>(r.V===NV.mi||r.V===NV.noc)&&r.AB===Y&&/^(6-12|1-3|Trên)/.test(r.ae)).reduce((a,r)=>a+U(r.L),0),invY),'%',+cfg.k3]];
 const TV={},TQ={};R.forEach(r=>{TV[r.u]=(TV[r.u]||0)+U(r.L);TQ[r.cd]=(TQ[r.cd]||0)+U(r.L)});
 const CVo={},CQo={};Object.keys(CV).forEach(c=>{if(TV[c]!==undefined)CVo[c]=[CV[c],TV[c]]});Object.keys(CQ).forEach(c=>{if(TQ[c]!==undefined)CQo[c]=[CQ[c],TQ[c],CQq[c]]});
 {const t=tabs.find(x=>x.n==='PL_HopDong');if(t&&Object.keys(CV).length){const n=t.a.length;let si=0,st=0;t.a.forEach((r,i)=>{if(!i)r.push('Giá trị HĐ ban đầu (nhập NCC)','Tồn hiện tại (giá vốn)','% tồn / ban đầu');else if(i===n-1)r.push(si,st,si?st/si*100:0);else{const o=CVo[r[0]];if(o){si+=o[0];st+=o[1];r.push(o[0],o[1],o[0]?o[1]/o[0]*100:0)}else r.push('','','')}})}}
 const BL=Object.values(bad).sort((a,b)=>b.v-a.v),top=o=>{let b,n=0;for(const x in o||{})if(o[x]>n){n=o[x];b=x}return b};BL.forEach(x=>{if(x.t==='nh')x.inf=INF[x.k];if(x.t==='lo')x.ct=x.e;if(x.t==='lk')x.sug=top(IL[x.e]);if(x.t==='nh')x.sug=top(IN[x.e])});
 const I=[],ad=(lv,k,t,d,need,go)=>I.push({lv,k,t,d,need,go}),f0=v=>Math.round(v).toLocaleString('en-US');
 [['td','Tiến độ lô hàng',1],['ncc','Nhập NCC',1],['tt','Tồn tiền'],['ex','Xuất kho'],['pnk','Nhập thu hồi'],['sn','Thu hồi chi tiết (serial)'],['tl','Xuất thanh lý']].filter(x=>!(S[x[0]]&&S[x[0]].rows.length)).forEach(x=>ad(x[2]?'err':'warn','m-'+x[0],'Chưa nạp '+x[1],'Các số liệu liên quan sẽ thiếu hoặc bằng 0.','Thả file '+x[1]+' vào ô nạp rồi bấm Tính toán lại.','nap'));
 const lo=BL.filter(x=>x.t==='lo');if(lo.length)ad('warn','lo',lo.length+' lô hàng chưa có ngày nhập ('+f0(lo.reduce((a,x)=>a+x.n,0))+' dòng, '+f0(lo.reduce((a,x)=>a+x.v,0)/rate)+' USD)','Năm nhập, hạn tồn và phân loại năm của các dòng này chưa xác định.'+(LI.size?' (Đã lấy ngày nhập từ Import_Summary theo mã HĐ cho '+LI.size+' lô khác.)':''),'Nạp Tiến độ lô hàng đủ kỳ (ngày lấy ở cột N, không có thì các cột ngày khác hoặc cột F) và Import_Summary NCC (mã HĐ + ngày nhập). Lô đã có ngày sẽ được ghi nhớ cho các tháng sau.','nap');
 const cl=BL.filter(x=>x.t!=='lo');if(cl.length)ad('warn','cls',cl.length+' mục kho/mã hàng chưa phân loại','Giá trị đang nằm trong nhóm "Chưa phân loại" nên các bảng TH, PL chưa đủ.','Mở trang Chưa phân loại, dùng gợi ý hoặc gán loại.','cpl');
 const dw=ds.filter(r=>r[4]!=='Khớp'),tx=ds.reduce((a,r)=>a+r[1],0),ty=ds.reduce((a,r)=>a+r[2],0);if(dw.length)ad(Math.abs(tx-ty)>0.01*tx?'err':'warn','ds',dw.length+' mã kho lệch giữa tồn biểu đồ và tồn tiền','Tổng chênh ròng '+f0(tx-ty)+' USD (tồn biểu đồ '+f0(tx)+' / tồn tiền '+f0(ty)+'); cộng giá trị tuyệt đối '+f0(gross)+' USD, sau khi bù trừ các mã chuyển/xếp khác kho còn '+f0(remG)+' USD. Nguyên nhân chính các kho: '+(()=>{const c={};dsw.forEach(r=>{const k=String(r[6]).replace(/ \(.*$/,'');c[k]=(c[k]||0)+Math.abs(r[3])});return Object.entries(c).sort((a,b)=>b[1]-a[1]).slice(0,3).map(e=>e[0]+' '+f0(e[1])).join('; ')})()+'.','Mở trang Đối soát: có nguyên nhân từng kho, mã gây lệch, so sánh tháng trước và dự báo; tải file chi tiết theo mã.','ds');
 if(ne.n)ad('info','ne',f0(ne.n)+' dòng hàng thu hồi/thanh lý/chờ KCS không có phiếu hoặc serial thu hồi ('+f0(ne.v/rate)+' USD)','Đang tính vào nhóm năm trước ('+(Y-1)+'), không tính vào nhóm "Thu hồi".','Nạp đủ Nhập thu hồi và Thu hồi chi tiết từ đầu năm; tool tự cộng dồn qua các tháng.','nap');
 kpi.forEach((k,i)=>{if(k[3]&&k[1]>k[3])ad('warn','k'+i,k[0]+' = '+k[1].toFixed(2)+'% vượt mục tiêu '+k[3]+'%','Chỉ tiêu chưa đạt.','Xem các bảng PL liên quan để xác định nhóm hàng cần xử lý.','ph')});
 I.sort((a,b)=>'err warn info'.indexOf(a.lv)-'err warn info'.indexOf(b.lv));

 /* ===== kiểm tra chất lượng dữ liệu, so kỳ trước, biểu đồ cầu ===== */
 const MKP=(S.mem||{}).kp||{},cmk=mk(cut),pmk=Object.keys(MKP).filter(k=>k<cmk).sort().pop(),PK=pmk?MKP[pmk]:null;
 kpi.forEach((k,i)=>{k[4]=PK&&PK.v&&PK.v[i]!=null?PK.v[i]:null});
 const KPM={...MKP};KPM[cmk]={v:kpi.map(k=>k[1]),tot};{const ks=Object.keys(KPM).sort().slice(-14),o={};ks.forEach(k=>o[k]=KPM[k]);Object.keys(KPM).forEach(k=>{if(!(k in o))delete KPM[k]})}
 let bridge=null;
 if(PK&&PK.tot!=null){const mo=mm[cmk]||{x:0,tl:0,th:0},nn=NM[cmk]||0,pl=PK.tot+nn+(mo.th||0)-(mo.x||0)-(mo.tl||0),adj=tot-pl;
  bridge={from:pmk,to:cmk,steps:[['Tồn '+pmk,PK.tot,'t'],['Nhập NCC',nn,'+'],['Nhập thu hồi',mo.th||0,'+'],['Xuất ra trạm',-(mo.x||0),'-'],['Thanh lý',-(mo.tl||0),'-'],['Điều chỉnh / khác',adj,'a'],['Tồn '+cmk,tot,'t']]}}
 const dq=[],dqa=(ok,t,d)=>dq.push({ok,t,d});
 {const ts=a=>a[a.length-1][a[0].length-1],t1=ts(tabs.find(t=>t.n==='Ton_ChiNhanh').a),t2=ts(tabs.find(t=>t.n==='TH').a);
  dqa(Math.abs(t1-tot)<1&&Math.abs(t2-tot)<1?1:2,'Tổng các bảng bằng tổng tồn','Tồn chi nhánh '+n0(t1)+', TH '+n0(t2)+', tổng dòng '+n0(tot)+' USD');
  const rd=rowsOf('ct').filter(r=>{const k=String(r[c.k]).trim();return k&&k[0]!=='('}).length;dqa(rd===R.length?1:0,'Số dòng đọc bằng số dòng tính',n0(rd)+' dòng hợp lệ trong file, '+n0(R.length)+' dòng được tính');
  const sc={};let ds2=0;R.forEach(r=>{const k=sk(r.sn);});rowsOf('ct').forEach(r=>{const k=sk(r[c.sn]);if(k){if(sc[k])ds2++;else sc[k]=1}});dqa(ds2?0:1,'Serial lặp trong tồn biểu đồ',ds2?ds2+' dòng có serial đã xuất hiện ở dòng trước, cần kiểm tra trùng dòng':'Không có serial lặp');
  const lo2=BL.filter(x=>x.t==='lo');dqa(lo2.length?0:1,'Lô chưa có ngày nhập',lo2.length?lo2.length+' lô, '+n0(lo2.reduce((a,x)=>a+x.v,0)/rate)+' USD':'Mọi lô đều có ngày');
  const cl2=BL.filter(x=>x.t!=='lo'),cv2=cl2.reduce((a,x)=>a+x.v,0)/rate;dqa(cl2.length?0:1,'Chưa phân loại',cl2.length?cl2.length+' mục, '+n0(cv2)+' USD ('+(tot?(cv2/tot*100).toFixed(1):0)+'% tồn)':'Không còn mục nào');
  const miss=[['td','Tiến độ lô hàng',2],['ncc','Nhập NCC',2],['tt','Tồn tiền',0],['ex','Xuất kho',0],['pnk','Nhập thu hồi',0],['sn','Thu hồi chi tiết',0],['tl','Xuất thanh lý',0]].filter(x=>!(S[x[0]]&&S[x[0]].rows.length));
  dqa(miss.some(x=>x[2])?2:miss.length?0:1,miss.length?'Thiếu file: '+miss.map(x=>x[1]).join(', '):'Đủ file nguồn',miss.length?'Các số liệu liên quan thiếu hoặc bằng 0':'Tất cả file nguồn đã nạp');
  const fd=cfg.fd||{};if(fd.ct&&fd.tt){const dd=Math.round((new Date(fd.ct)-new Date(fd.tt))/DAY);dqa(Math.abs(dd)>3?0:1,'Ngày chốt hai báo cáo đối soát','Tồn biểu đồ '+fd.ct+', tồn tiền '+fd.tt+(Math.abs(dd)>3?': lệch '+Math.abs(dd)+' ngày, đối soát có thể lệch do chênh ngày':': khớp ngày'))}
  const ng=R.filter(r=>r.L<=0).length;dqa(ng?0:1,'Dòng giá trị âm hoặc bằng 0',ng?ng+' dòng, nên kiểm tra':'Không có');
  dqa(ds.filter(r=>r[4]!=='Khớp').length?0:1,'Đối soát tồn biểu đồ và tồn tiền',ds.filter(r=>r[4]!=='Khớp').length+' / '+ds.length+' mã kho lệch')}
 const dscore=Math.max(0,100-dq.reduce((a,x)=>a+(x.ok===2?15:x.ok===0?6:0),0));
 const files={have:['cat','ct','tt','td','ncc','pnk','sn','ex','tl'].filter(k=>k==='cat'?!!S.cat:S[k]&&S[k].rows&&S[k].rows.length).length,total:9};
 return{dq:{score:dscore,checks:dq,files},bridge,inv,invY,cv:CVo,cq:CQo,tabs,bad:BL,kpi,n:R.length,issues:I,rows:R,mem:{kp:KPM,ds:S.__ds,ic:IC,lo:SH,pn:[...PN],sn:[...SN],mon:mm}}};
})();
if(typeof module!=='undefined')module.exports=NT;
