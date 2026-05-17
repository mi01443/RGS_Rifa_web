// ============================================================
//  RifaFácil — Google Apps Script
//  IMPORTANTE: Usa apenas GET para evitar o problema de
//  redirect 302 que perde o body do POST.
//  Todas as operações chegam via e.parameter.
// ============================================================

var ABA_CONFIG   = 'Config';
var ABA_PIX      = 'Pix';
var ABA_NUMEROS  = 'Numeros';
var ABA_SETTINGS = 'Settings';

// Índices das colunas (0-based)
var C_NUM = 0, C_NOME = 1, C_TEL = 2, C_STATUS = 3,
    C_DATA = 4, C_RESATE = 5, C_PID = 6;

// ── Setup ─────────────────────────────────────────────────────
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  function mkSheet(name, headers) {
    var s = ss.getSheetByName(name) || ss.insertSheet(name);
    s.getRange(1,1,1,headers.length).setValues([headers])
     .setBackground('#1a1a2e').setFontColor('#f5c542').setFontWeight('bold');
    return s;
  }
  mkSheet(ABA_CONFIG,   ['chave','valor']);
  mkSheet(ABA_PIX,      ['chave','valor']);
  mkSheet(ABA_SETTINGS, ['chave','valor']);
  mkSheet(ABA_NUMEROS,  ['numero','nome','telefone','status','data','reservado_ate','pedido_id']);

  var cfg = ss.getSheetByName(ABA_CONFIG);
  if (cfg.getLastRow() <= 1)
    cfg.getRange(2,1,7,2).setValues([
      ['title','Minha Rifa'],['desc','Clique em um número para participar!'],
      ['price','10'],['qty','100'],['date',''],['imgUrl',''],['password','admin123']
    ]);

  var pix = ss.getSheetByName(ABA_PIX);
  if (pix.getLastRow() <= 1)
    pix.getRange(2,1,4,2).setValues([['type','CPF'],['key',''],['name','Organizador'],['city','Brasil']]);

  SpreadsheetApp.getUi().alert('✅ RifaFácil configurado!\n\nAgora faça o Deploy como Web App.');
}

// ── Helpers ───────────────────────────────────────────────────
function toInt(v) {
  if (v===''||v===null||v===undefined) return NaN;
  return parseInt(String(v).replace(/\.0+$/,'').trim(),10);
}

function sheetToObj(name) {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!s) return {};
  var rows = s.getDataRange().getValues(), obj = {};
  for (var i=1;i<rows.length;i++) {
    var k=rows[i][0];
    if (k===''||k===null||k===undefined) continue;
    var v=rows[i][1];
    if (v instanceof Date) {
      v = v.getFullYear()+'-'+String(v.getMonth()+1).padStart(2,'0')+'-'+String(v.getDate()).padStart(2,'0');
    }
    obj[String(k)] = (v===null||v===undefined)?'':String(v);
  }
  return obj;
}

function objToSheet(name, obj) {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!s) return;
  var last=s.getLastRow();
  if (last>1) s.getRange(2,1,last-1,2).clearContent();
  var keys=Object.keys(obj);
  if (keys.length>0) s.getRange(2,1,keys.length,2).setValues(keys.map(function(k){return [k,obj[k]];}));
}

function delRows(sheet, idxs) {
  idxs.sort(function(a,b){return b-a;}).forEach(function(r){sheet.deleteRow(r);});
}

function makePid() {
  var c='ABCDEFGHJKLMNPQRSTUVWXYZ23456789', r='';
  for(var i=0;i<5;i++) r+=c[Math.floor(Math.random()*c.length)];
  return String(new Date().getTime()).slice(-5)+r;
}

function getNumeros() {
  var s=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_NUMEROS);
  if (!s||s.getLastRow()<=1) return {};
  var rows=s.getDataRange().getValues(), now=new Date().getTime(), obj={}, del=[];
  for (var i=1;i<rows.length;i++) {
    var num=toInt(rows[i][C_NUM]); if(isNaN(num)) continue;
    var st=String(rows[i][C_STATUS]||'').trim().toLowerCase();
    var ra=rows[i][C_RESATE];
    var exp=ra?(ra instanceof Date?ra.getTime():new Date(String(ra)).getTime()):0;
    if (st==='cancelled'||st==='expired') { del.push(i+1); continue; }
    if (st==='reserved'&&exp&&exp<now)    { del.push(i+1); continue; }
    var dv=rows[i][C_DATA];
    obj[String(num)]={
      name:   String(rows[i][C_NOME]||''),
      phone:  String(rows[i][C_TEL]||''),
      status: st,
      date:   dv instanceof Date?dv.toISOString():String(dv||''),
      reservedUntil: ra?(ra instanceof Date?ra.toISOString():String(ra)):'',
      pedidoId: String(rows[i][C_PID]||''),
    };
  }
  if(del.length) delRows(s,del);
  return obj;
}

function out(p) {
  return ContentService.createTextOutput(JSON.stringify(p))
    .setMimeType(ContentService.MimeType.JSON);
}
function ok(d)   { return out({ok:true, data:d}); }
function fail(m) { return out({ok:false,error:m}); }

// ── Único ponto de entrada: doGet ─────────────────────────────
// TUDO chega via GET. Operações de leitura usam ?action=...
// Operações de escrita chegam via ?action=...&p0=...&p1=... etc.
function doGet(e) {
  try {
    var p      = e && e.parameter ? e.parameter : {};
    var action = p.action || 'all';

    // ── LEITURAS ─────────────────────────────────────────────
    if (action==='all')      return ok({config:sheetToObj(ABA_CONFIG),pix:sheetToObj(ABA_PIX),settings:sheetToObj(ABA_SETTINGS),numbers:getNumeros()});
    if (action==='numbers')  return ok({numbers:getNumeros()});
    if (action==='config')   return ok({config:sheetToObj(ABA_CONFIG)});
    if (action==='settings') return ok({settings:sheetToObj(ABA_SETTINGS)});

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ── SALVAR CONFIG / PIX / SETTINGS ───────────────────────
    if (action==='saveConfig')   { objToSheet(ABA_CONFIG,  JSON.parse(p.data)); return ok('ok'); }
    if (action==='savePix')      { objToSheet(ABA_PIX,     JSON.parse(p.data)); return ok('ok'); }
    if (action==='saveSettings') { objToSheet(ABA_SETTINGS,JSON.parse(p.data)); return ok('ok'); }

    // ── RESERVAR ─────────────────────────────────────────────
    if (action==='reserveNumbers') {
      var sheet  = ss.getSheetByName(ABA_NUMEROS);
      var nums   = JSON.parse(p.numbers);
      var nome   = String(p.name||'').trim();
      var tel    = String(p.phone||'').trim();
      var now    = new Date();
      var expiry = new Date(now.getTime()+15*60*1000).toISOString();
      var pid    = makePid();
      var atual  = getNumeros();
      var confl  = nums.filter(function(n){return atual[String(parseInt(n,10))];});
      if (confl.length>0) return ok({conflicts:confl});
      var rows=nums.map(function(n){return [parseInt(n,10),nome,tel,'reserved',now.toISOString(),expiry,pid];});
      sheet.getRange(sheet.getLastRow()+1,1,rows.length,7).setValues(rows);
      return ok({conflicts:[],expiresAt:expiry,pedidoId:pid});
    }

    // ── CONFIRMAR COMPRA (reserved → pending) ─────────────────
    if (action==='confirmNumbers') {
      var sheet  = ss.getSheetByName(ABA_NUMEROS);
      var pid    = String(p.pedidoId||'');
      var rows   = sheet.getDataRange().getValues();
      var count  = 0;
      for (var i=1;i<rows.length;i++) {
        if (String(rows[i][C_PID]||'')===pid && String(rows[i][C_STATUS]||'').trim().toLowerCase()==='reserved') {
          sheet.getRange(i+1,C_STATUS+1).setValue('pending');
          sheet.getRange(i+1,C_RESATE+1).setValue('');
          count++;
        }
      }
      return ok({confirmed:count});
    }

    // ── CANCELAR RESERVA ──────────────────────────────────────
    if (action==='cancelReservation') {
      var sheet = ss.getSheetByName(ABA_NUMEROS);
      var pid   = String(p.pedidoId||'');
      var rows  = sheet.getDataRange().getValues(), del=[];
      for (var i=1;i<rows.length;i++)
        if (String(rows[i][C_PID]||'')===pid) del.push(i+1);
      delRows(sheet,del);
      return ok('cancelado');
    }

    // ── ADMIN: CONFIRMAR PAGAMENTO (pending → paid) ───────────
    if (action==='confirmPedido') {
      var sheet = ss.getSheetByName(ABA_NUMEROS);
      var pid   = String(p.pedidoId||'');
      var rows  = sheet.getDataRange().getValues();
      var count = 0;
      for (var i=1;i<rows.length;i++) {
        if (String(rows[i][C_PID]||'')===pid) {
          var st=String(rows[i][C_STATUS]||'').trim().toLowerCase();
          if (st==='pending'||st==='reserved') {
            sheet.getRange(i+1,C_STATUS+1).setValue('paid');
            sheet.getRange(i+1,C_RESATE+1).setValue('');
            count++;
          }
        }
      }
      return ok({confirmed:count, pedidoId:pid});
    }

    // ── ADMIN: CANCELAR PEDIDO ────────────────────────────────
    if (action==='cancelPedido') {
      var sheet = ss.getSheetByName(ABA_NUMEROS);
      var pid   = String(p.pedidoId||'');
      var rows  = sheet.getDataRange().getValues(), del=[];
      for (var i=1;i<rows.length;i++)
        if (String(rows[i][C_PID]||'')===pid) del.push(i+1);
      delRows(sheet,del);
      return ok('cancelado');
    }

    // ── ADMIN: REVERTER PAGAMENTO (paid → pending) ────────────
    if (action==='revertPedido') {
      var sheet = ss.getSheetByName(ABA_NUMEROS);
      var pid   = String(p.pedidoId||'');
      var rows  = sheet.getDataRange().getValues();
      for (var i=1;i<rows.length;i++)
        if (String(rows[i][C_PID]||'')===pid)
          sheet.getRange(i+1,C_STATUS+1).setValue('pending');
      return ok('revertido');
    }

    // ── ZERAR TUDO ────────────────────────────────────────────
    if (action==='resetNumbers') {
      var sheet=ss.getSheetByName(ABA_NUMEROS), last=sheet.getLastRow();
      if (last>1) sheet.deleteRows(2,last-1);
      return ok('zerado');
    }

    // ── UPLOAD DE IMAGEM (ainda via POST pois base64 é grande) ─
    return fail('Ação desconhecida: '+action);
  } catch(err) {
    return fail('Erro: '+err.message+' | Stack: '+(err.stack||''));
  }
}

// POST só para upload de imagem (base64 é grande demais para URL)
function doPost(e) {
  try {
    var body   = JSON.parse(e.postData.contents);
    if (body.action !== 'uploadImage') return fail('POST só aceita uploadImage');
    var raw    = body.base64.replace(/^data:image\/\w+;base64,/,'');
    var ext    = (body.ext||'jpg').toLowerCase();
    var mimes  = {jpg:'image/jpeg',jpeg:'image/jpeg',png:'image/png',gif:'image/gif',webp:'image/webp'};
    var blob   = Utilities.newBlob(Utilities.base64Decode(raw),mimes[ext]||'image/jpeg','prize.'+ext);
    var folder = (function(nm){var f=DriveApp.getFoldersByName(nm);return f.hasNext()?f.next():DriveApp.createFolder(nm);})('RifaFacil_Imagens');
    var ex=folder.getFilesByName('prize.'+ext);
    while(ex.hasNext()) ex.next().setTrashed(true);
    var file=folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);
    return out({ok:true,data:{url:'https://drive.google.com/thumbnail?id='+file.getId()+'&sz=w800'}});
  } catch(err) { return out({ok:false,error:'Erro upload: '+err.message}); }
}
