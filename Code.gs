/** CENTRAL DE COLETAS MM — Google Apps Script web app.
 * A planilha oficial permanece a única fonte dos pedidos brutos.
 * Implantação: executar como proprietário; acesso apenas a colaboradores autorizados.
 */
const MM_CONFIG = Object.freeze({
  SPREADSHEET_ID: '1SpVBGjAVLRlMgAxCDmE2-9x_IppO0cpk1JVROp8g_7g',
  TIMEZONE: 'America/Sao_Paulo',
  ACCOUNTS: ['AC', 'MP', 'MM', 'RAFRISTAL'],
  PLATFORMS: ['Shopee', 'TikTok Shop', 'Mercado Livre', 'Shein'],
});

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Central de Coletas MM')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getDashboard(request) {
  request = request || {};
  const dateString = String(request.date || Utilities.formatDate(new Date(), MM_CONFIG.TIMEZONE, 'yyyy-MM-dd'));
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateString)) throw Error('Informe uma data no formato AAAA-MM-DD.');
  const [year,month,day] = dateString.split('-').map(Number);
  const d = new Date(year,month-1,day,12,0,0);
  if (d.getFullYear()!==year || d.getMonth()!==month-1 || d.getDate()!==day) throw Error('Data inválida.');
  const tabName = ['JANEIRO','FEVEREIRO','MARÇO','ABRIL','MAIO','JUNHO','JULHO','AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO'][month-1] + ' ' + year;
  // A planilha e lida novamente a CADA consulta; nenhum total e mantido em cache.\n  const ss = SpreadsheetApp.openById(MM_CONFIG.SPREADSHEET_ID);
  const tab = ss.getSheetByName(tabName);
  if (!tab) throw Error('Aba não encontrada na planilha oficial: ' + tabName);
  const last = tab.getLastRow();
  if (last < 1) return {date:dateString,tab:tabName,rows:[],totals:{},range:null,loadedAt:new Date().toISOString()};
  const data = tab.getRange(1,1,last,6).getValues();
  const DISPLAY = tab.getRange(1,1,last,6).getDisplayValues();
  const key = Utilities.formatDate(d, MM_CONFIG.TIMEZONE, 'dd/MM/yyyy');
  const dateHeading = /SEGUNDA|TERÇA|TERCA|QUARTA|QUINTA|SEXTA|SABADO|SÁBADO/i;
  function normalizedDate(cell,rendered) {
    if (cell instanceof Date && !isNaN(cell)) return Utilities.formatDate(cell,MM_CONFIG.TIMEZONE,'dd/MM/yyyy');
    let s = String(rendered || cell || '').trim();
    if (/^\d{1,2}[\/-]\d{1,2}[\/-]\d{4}$/.test(s)) {
      const v=s.split(/[\/-]/);return v[0].padStart(2,'0')+'/'+v[1].padStart(2,'0')+'/'+v[2];
    }
    return s;
  }
  let start=-1, end=data.length;
  for(let i=0;i<data.length;i++) {
    const heading=String(DISPLAY[i][4]||'');
    if (!dateHeading.test(heading)) continue;
    if (start>=0) {end=i;break;}
    if (normalizedDate(data[i][5],DISPLAY[i][5])===key) start=i;
  }
  if (start<0) throw Error('Não encontrei o bloco da data '+key+' na aba '+tabName+'.');
  // Coluna C é propositalmente ignorada, salvo para identificar o CABEÇALHO de ML PRÓXIMOS DIAS.
  let section=''; const rows=[];const totals={};
  for(let i=start+1;i<end;i++) {
    const b=String(DISPLAY[i][1]||'').trim().toUpperCase().replace(/\s+/g,' ');
    const c=String(DISPLAY[i][2]||'').toUpperCase();
    if(b.includes('SHOPEE DIA ANTERIOR')) {section='SHOPEE_1';continue;}
    if(b.includes('SHOPEE DO DIA')) {section='SHOPEE_2';continue;}
    if(b==='TIK TOK'||b==='TIKTOK'||b==='TIKTOK SHOP') {section='TIKTOK';continue;}
    if(b==='SHEIN') {section='SHEIN';continue;}
    if(b==='ML' && c.includes('PROXIMOS DIAS')) {section='ML_FUTURO';continue;}
    if(b==='ML') {section='ML';continue;}
    if(!['SHOPEE_1','SHOPEE_2','TIKTOK','SHEIN','ML'].includes(section)) continue;
    let account=b;
    if (account==='TK - MM') account='MM';
    if (account==='SH-MP') account='MP';
    if (!MM_CONFIG.ACCOUNTS.includes(account)) continue;
    const raw=data[i][3];
    let amount=typeof raw==='number'?raw:Number(String(DISPLAY[i][3]||'').replace(/\./g,'').replace(',','.'));
    if (!Number.isFinite(amount)||amount<=0) continue;
    if (!Number.isInteger(amount)) throw Error('Quantidade não inteira na linha '+(i+1));
    const platform=section.startsWith('SHOPEE')?'Shopee':section==='TIKTOK'?'TikTok Shop':section==='ML'?'Mercado Livre':'Shein';
    const stage=section==='SHOPEE_1'?'Primeira coleta':section==='SHOPEE_2'?'Segunda coleta':'';
    const k=platform+'|'+stage+'|'+account;
    totals[k]=(totals[k]||0)+amount;
    rows.push({row:i+1,platform,stage,account,amount});
  }
  return {date:dateString,dateBR:key,tab:tabName,range:{start:start+1,end},rows,totals,recordCount:rows.length,loadedAt:Utilities.formatDate(new Date(),MM_CONFIG.TIMEZONE,'dd/MM/yyyy HH:mm:ss'),sourceUrl:ss.getUrl()};
}
