const {test} = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');

const source = fs.readFileSync('apps-script/Code.gs', 'utf8');
const id = '11111111-1111-4111-8111-111111111111';
const valid = {name:'測試姓名', organization:'測試醫院', job:'藥師', profession:'藥事',
  host:'是', attendance:'線上', email:'test@example.com', phone:'0912345678', requestId:id};
const expectedHeaders = fs.readFileSync('apps-script/headers.tsv', 'utf8').trim().split('\t');

function harness(options = {}) {
  const rows = options.rows || [];
  const state = {rows, formats:[], acquired:0, released:0, opens:[], inserted:[], flushes:0};
  const sheet = {
    getMaxColumns:()=>10, getMaxRows:()=>1000, getLastRow:()=>rows.length,
    insertColumnsAfter(){}, insertRowsAfter(){}, setFrozenRows(){}, autoResizeColumns(){},
    getRange(row, column, count=1, columns=1) {
      const range = {
        setValues(values) {values.forEach((values,i)=>{rows[row-1+i] ||= [];
          values.forEach((value,j)=>{rows[row-1+i][column-1+j]=value;});});return range;},
        getValues() {return Array.from({length:count},(_,i)=>Array.from({length:columns},(_,j)=>rows[row-1+i]?.[column-1+j] ?? ''));},
        setNumberFormat(value) {state.formats.push({row,column,count,columns,value});return range;},
        setFontWeight(){return range;}, setBackground(){return range;}, setFontColor(){return range;},
        createTextFinder(value) {let exact=false;return {
          matchEntireCell(flag){exact=flag;return this;},
          findNext(){assert.equal(exact,true);return rows.slice(row-1,row-1+count).some(r=>r[column-1]===value)?{}:null;}
        };}
      };return range;
    }
  };
  const lock = {held:false,waitLock(){this.held=true;state.acquired++;},
    tryLock(){if(options.lockDenied)return false;this.held=true;state.acquired++;return true;},
    hasLock(){return this.held;},releaseLock(){this.held=false;state.released++;}};
  const spreadsheet={getSheetByName(name){assert.equal(name,'工作坊報名資料');return options.missingSheet ? null : sheet;},
    insertSheet(name){state.inserted.push(name);return sheet;}};
  const context = vm.createContext({console:{log(){}},LockService:{getScriptLock:()=>lock},
    SpreadsheetApp:{openById(value){state.opens.push(value);if(options.openDenied)throw Error('access denied');return spreadsheet;},
      flush(){state.flushes++;if(options.failFirstFlush && state.flushes===1)throw Error('write confirmation failed');}},
    Utilities:{formatDate(date,tz,format){assert.equal(tz,'Asia/Taipei');assert.equal(format,'yyyy/MM/dd HH:mm:ss');return '2026/10/09 10:30:00';}},
    HtmlService:{createHtmlOutput(html){return {html,setTitle(){return this;},addMetaTag(){return this;}};}}
  });
  vm.runInContext(source,context);
  const post=(overrides={})=>context.doPost({parameter:{...valid,...overrides},postData:{length:500}}).html;
  return {context,state,post};
}

test('setup creates the named sheet and exact TSV headers without deleting data',()=>{
  const {context,state}=harness({missingSheet:true});context.setupSheet();
  assert.deepEqual(state.inserted,['工作坊報名資料']);
  assert.deepEqual(Array.from(state.rows[0]),expectedHeaders);
  assert.equal(state.opens[0],'1LEllNReIM9pXQ6M3Jac7CgOgfJ29gNWz5t9qlcP0D1Y');
  assert.equal(state.acquired,state.released);
});

test('valid registration preserves field order, leading-zero phone and server timestamp',()=>{
  const {post,state}=harness();assert.match(post(),/報名已送出/);
  assert.deepEqual(Array.from(state.rows[1]),['2026/10/09 10:30:00',valid.name,valid.organization,
    valid.job,valid.profession,valid.host,valid.attendance,valid.email,valid.phone,id]);
  assert.equal(state.formats[0].value,'@');assert.equal(state.flushes,1);assert.equal(state.released,1);
});

test('same ID retried after write acknowledgement failure creates exactly one row',()=>{
  const {post,state}=harness({failFirstFlush:true});assert.match(post(),/報名尚未確認/);
  assert.match(post(),/這筆報名已收到/);assert.equal(state.rows.length,2);
  assert.equal(state.released,2);
});

test('different registration ID creates a new row even with same email',()=>{
  const {post,state}=harness();post();post({requestId:'22222222-2222-4222-8222-222222222222'});
  assert.equal(state.rows.length,3);
});

test('invalid required fields, lengths, enums, email, phone or UUID never touch spreadsheet',()=>{
  for(const overrides of [{name:' '},{name:'a'.repeat(81)},{profession:'unknown'},
    {host:'maybe'},{attendance:'混合'},{email:'not-an-email'},{phone:'abc12345'},
    {phone:'123'},{requestId:'no-id'}]) {
    const {post,state}=harness();assert.match(post(overrides),/報名尚未確認/);
    assert.equal(state.rows.length,0);assert.equal(state.opens.length,0);
  }
});

test('missing, oversized and repeated-field requests are rejected before writes',()=>{
  for(const event of [undefined,{parameter:valid,postData:{length:10001}},
    {parameter:valid,postData:{length:500},parameters:{name:['first','second']}}]){
    const {context,state}=harness();assert.match(context.doPost(event).html,/報名尚未確認/);
    assert.equal(state.opens.length,0);
  }
});

test('mismatched existing header preserves every existing cell',()=>{
  const rows=[['姓名','Email'],['原有姓名','original@example.com']];const snapshot=JSON.stringify(rows);
  const {post,state}=harness({rows});assert.match(post(),/表頭與程式不一致/);
  assert.equal(JSON.stringify(state.rows),snapshot);assert.equal(state.released,1);
});

test('formula-shaped fields are escaped and international phones remain text',()=>{
  const {post,state}=harness();post({organization:'=IMPORTXML("url","path")',job:'@SUM(1)',phone:'+886912345678'});
  assert.equal(state.rows[1][2],"'=IMPORTXML(\"url\",\"path\")");
  assert.equal(state.rows[1][3],"'@SUM(1)");assert.equal(state.rows[1][8],"'+886912345678");
});

test('lock contention reports unconfirmed and writes nothing',()=>{
  const {post,state}=harness({lockDenied:true});assert.match(post(),/尚未確認資料寫入/);
  assert.equal(state.rows.length,0);assert.equal(state.opens.length,0);
});

test('service permission failure releases lock and does not report success',()=>{
  const {post,state}=harness({openDenied:true});const html=post();
  assert.match(html,/報名尚未確認/);assert.doesNotMatch(html,/報名已送出|access denied/);
  assert.equal(state.released,1);
});

test('GET and receipts do not expose submitted names, email or phone; HTML is escaped',()=>{
  const {post,context}=harness();const html=post({name:'<script>alert(1)</script>'});
  assert.doesNotMatch(html,/<script>|test@example.com|0912345678/);
  assert.doesNotMatch(context.doGet().html,/1LEllNRe|test@example.com/);
  assert.match(context.receipt_('<img src=x>','<script>bad</script>',false).html,/&lt;img/);
  assert.doesNotMatch(context.receipt_('x','<script>bad</script>',false).html,/<script>/);
});
