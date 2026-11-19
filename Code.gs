/**
 * ABPS – Google Apps Script (Code.gs)
 * प्रथम एकदा setup() फंक्शन Run करा – ४ टॅब आपोआप तयार होतील:
 *   GP_List       : A=जिल्हा, B=तालुका, C=ग्रामपंचायत (English), D=ग्रामपंचायत (मराठी), E=गाव (एका गावासाठी एक ओळ)
 *   Banks         : A=बँकेचा प्रकार (Commercial Bank / Co-operative Bank / Post Office), B=बँकेचे नाव
 *   Aadhaar Data  : नोंदी (आपोआप)
 *   Bank Data     : नोंदी (आपोआप). शेवटचा लपवलेला स्तंभ "Village (helper)" – गावानुसार शोध/डाउनलोडसाठी; Excel मध्ये येत नाही.
 *
 * एकच नोंद = Panchayat + Village + Family ID + Applicant No. ती पुन्हा आली तर दुबार होत नाही;
 * माहिती बदलली असेल तर पुष्टीनंतर जुनी नोंद बदलली जाते (replace).
 */
const GH = ['Jilha','Taluka','GP_English','GP_Marathi','Village'];
const BKH = ['Bank Type','Bank Name'];
const AH = ['Panchayat','Village','Family ID','Applicant No','Name','Aadhaar No','Name as per Aadhaar'];
const BH = ['Panchayat','Family ID','Applicant No','Name','Bank','IFSC','Account No','Account Holder Name','Village (helper)'];

function out_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
const s_ = x => String(x == null ? '' : x).trim();
const n_ = x => s_(x).toLowerCase();

function sh_(name, head){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if(!sh){
    sh = ss.insertSheet(name);
    sh.appendRow(head);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,head.length).setFontWeight('bold').setBackground('#d9ead3');
  }
  if(name === 'Bank Data' && s_(sh.getRange(1,9).getValue()) !== 'Village (helper)'){
    sh.getRange(1,9).setValue('Village (helper)').setFontWeight('bold');
  }
  if(name === 'Bank Data' && !sh.isColumnHiddenByUser(9)) sh.hideColumns(9);
  return sh;
}

function setup(){
  sh_('Aadhaar Data', AH); sh_('Bank Data', BH);
  const g = sh_('GP_List', GH);
  if(g.getLastRow() < 2) g.appendRow(['Satara','Satara','Sample GP','नमुना ग्रामपंचायत','Sample Village']);
  const b = sh_('Banks', BKH);
  if(b.getLastRow() < 2){
    [['Commercial Bank','State Bank of India'],['Commercial Bank','Bank of Maharashtra'],['Commercial Bank','Bank of India'],
     ['Commercial Bank','Union Bank of India'],['Commercial Bank','Canara Bank'],['Commercial Bank','Bank of Baroda'],
     ['Commercial Bank','HDFC Bank'],['Commercial Bank','ICICI Bank'],['Commercial Bank','Axis Bank'],
     ['Co-operative Bank','Satara District Central Co-operative Bank'],
     ['Post Office','India Post Payments Bank'],['Post Office','Post Office Savings Bank']
    ].forEach(r => b.appendRow(r));
  }
}

function doGet(e){
  const p = e.parameter, a = p.action || 'init';
  if(a === 'init'){
    const gp = sh_('GP_List', GH).getDataRange().getValues().slice(1).filter(r => r[0] && r[2])
      .map(r => ({j:s_(r[0]), t:s_(r[1]), e:s_(r[2]), m:s_(r[3]), v:s_(r[4])}));
    const bk = sh_('Banks', BKH).getDataRange().getValues().slice(1).filter(r => r[0] && r[1])
      .map(r => ({t:s_(r[0]), n:s_(r[1])}));
    return out_({gp:gp, banks:bk});
  }
  if(a === 'data'){
    const en = new Set(), mr = new Set();
    sh_('GP_List', GH).getDataRange().getValues().slice(1).forEach(r => {
      if(s_(r[0]) === p.jilha && s_(r[1]) === p.taluka && (!p.gp || s_(r[2]) === p.gp)){ en.add(s_(r[2])); mr.add(s_(r[3])); }
    });
    const vil = n_(p.village || '');
    const av = sh_('Aadhaar Data', AH).getDataRange().getDisplayValues();
    const bv = sh_('Bank Data', BH).getDataRange().getDisplayValues();
    const A = [av[0]].concat(av.slice(1).filter(r => en.has(s_(r[0])) && (!vil || n_(r[1]) === vil)));
    const B = [bv[0].slice(0,8)].concat(bv.slice(1).filter(r => mr.has(s_(r[0])) && (!vil || n_(r[8]) === vil)).map(r => r.slice(0,8)));
    return out_({a:A, b:B});
  }
  return out_({ok:false});
}

function put_(sh, idx, row){
  const r = idx >= 0 ? idx + 1 : sh.getLastRow() + 1;
  sh.getRange(r,1,1,row.length).setNumberFormat('@').setValues([row.map(String)]);
}

function doPost(e){
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try{
    const d = JSON.parse(e.postData.contents);
    const A = sh_('Aadhaar Data', AH), B = sh_('Bank Data', BH);
    const av = A.getDataRange().getDisplayValues(), bv = B.getDataRange().getDisplayValues();
    let ai = -1, bi = -1;
    for(let i = 1; i < av.length; i++){ const r = av[i];
      if(n_(r[0]) === n_(d.gpE) && n_(r[1]) === n_(d.village) && n_(r[2]) === n_(d.family) && n_(r[3]) === n_(d.applicant)){ ai = i; break; } }
    for(let i = 1; i < bv.length; i++){ const r = bv[i];
      if(n_(r[0]) === n_(d.gpM) && n_(r[1]) === n_(d.family) && n_(r[2]) === n_(d.applicant) && n_(r[8]) === n_(d.village)){ bi = i; break; } }

    const rowA = [d.gpE, d.village, d.family, d.applicant, d.name, d.aadhaar, d.aadhaarName];
    const rowB = [d.gpM, d.family, d.applicant, d.name, d.bank, d.ifsc, d.account, d.accountName, d.village];

    if(ai < 0 && bi < 0){ put_(A,-1,rowA); put_(B,-1,rowB); return out_({ok:true, status:'new'}); }

    const oa = ai >= 0 ? av[ai] : null, ob = bi >= 0 ? bv[bi] : null;
    const old = {
      name: s_(oa ? oa[4] : ob[3]), aadhaar: s_(oa ? oa[5] : ''), aadhaarName: s_(oa ? oa[6] : ''),
      bank: s_(ob ? ob[4] : ''), ifsc: s_(ob ? ob[5] : ''), account: s_(ob ? ob[6] : ''), accountName: s_(ob ? ob[7] : '')
    };
    const changed = Object.keys(old).filter(k => old[k] !== s_(d[k]));
    if(!changed.length) return out_({ok:true, status:'same'});
    if(!d.replace) return out_({ok:false, status:'exists', old:old, changed:changed});

    put_(A, ai, rowA); put_(B, bi, rowB);
    return out_({ok:true, status:'replaced'});
  } catch(err){
    return out_({ok:false, error:String(err)});
  } finally { lock.releaseLock(); }
}
