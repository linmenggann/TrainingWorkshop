const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
const headers=fs.readFileSync('apps-script/headers.tsv','utf8').trim().split('\t');
const source=fs.readFileSync('apps-script/Code.gs','utf8');
const fixture=()=>[
 headers.slice(),
 ['2026/10/08 09:00:00','測試甲','機構甲','藥師','藥事','是','實體','a@example.com','0911111111','id-a'],
 ['2026/10/09 11:00:00','測試乙','機構甲','護理師','護理','否','線上','b@example.com','0922222222','id-b'],
 [new Date('2026-10-08T17:00:00Z'),'測試丙','機構乙','藥師','藥事','是','線上','c@example.com','0933333333','id-c'],
 ['2026/02/30 08:00:00','測試丁','機構丙','教學','其他','未知','混合','d@example.com','0944444444','id-d'],
 ['', '測試戊','','教學','','否','實體','','','id-e'],
 Array(10).fill('')
];
function harness(rows=fixture(),options={}){
 const state={reads:0,writes:0};
 const sheet={getLastRow:()=>rows.length,getMaxColumns:()=>10,
  getRange(row,col,count,columns){state.reads++;return {getValues:()=>rows.slice(row-1,row-1+count).map(r=>r.slice(col-1,col-1+columns)),
   setValues(){state.writes++;throw Error('Stats must not write');}}}};
 class FixedDate extends Date{constructor(...args){super(...(args.length?args:['2026-10-08T17:30:00Z']));}}
 const context=vm.createContext({Date:FixedDate,
  SpreadsheetApp:{openById(id){assert.equal(id,'1LEllNReIM9pXQ6M3Jac7CgOgfJ29gNWz5t9qlcP0D1Y');if(options.accessDenied)throw Error('private error');return {
   getSheetByName(name){assert.equal(name,'工作坊報名資料');return options.missing?null:sheet;},
   insertSheet(){state.writes++;throw Error('Must not create sheet');}};}},
  Utilities:{formatDate(date,tz,format){assert.equal(tz,'Asia/Taipei');
   const p=Object.fromEntries(new Intl.DateTimeFormat('en-CA',{timeZone:tz,year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',second:'2-digit',hourCycle:'h23'}).formatToParts(date).map(p=>[p.type,p.value]));
   if(format==='yyyy-MM-dd')return `${p.year}-${p.month}-${p.day}`;
   assert.equal(format,'yyyy/MM/dd HH:mm:ss');return `${p.year}/${p.month}/${p.day} ${p.hour}:${p.minute}:${p.second}`;}},
  ContentService:{MimeType:{JSON:'json',JAVASCRIPT:'javascript'},createTextOutput(body){return {body,setMimeType(type){this.type=type;return this;}};}}
 });vm.runInContext(source,context);
 const get=params=>context.doGet({parameter:{action:'dashboard',...params}});
 const stats=params=>JSON.parse(get(params).body);
 return {context,state,get,stats};
}

test('aggregate totals, distinct institutions, Taipei dates and distributions match data rows',()=>{
 const {stats,state}=harness();const data=stats();assert.equal(data.ok,true);assert.equal(data.total,5);assert.equal(data.overallTotal,5);
 assert.equal(data.organizations,3);assert.equal(data.today,2);assert.equal(data.professions['藥事'],2);
 assert.deepEqual(data.attendance,{'實體':2,'線上':2,'未分類':1});assert.deepEqual(data.hosts,{'是':2,'否':2,'未分類':1});
 assert.deepEqual(data.daily,[{date:'2026-10-08',count:1},{date:'2026-10-09',count:2}]);
 assert.deepEqual(data.quality,{incomplete:1,unknownDate:2});assert.equal(data.lastRegistrationAt,'2026-10-09 11:00:00');
 assert.equal(state.writes,0);
});
test('inclusive date filters exclude undated rows and compute institutions within filter',()=>{
 const data=harness().stats({from:'2026-10-09',to:'2026-10-09'});
 assert.equal(data.total,2);assert.equal(data.overallTotal,5);assert.equal(data.organizations,2);assert.equal(data.today,2);
 assert.equal(data.quality.unknownDate,0);assert.equal(data.daily.length,1);
});
test('combined profession, attendance and host filters are applied before charts',()=>{
 const data=harness().stats({profession:'藥事',attendance:'線上',host:'是'});
 assert.equal(data.total,1);assert.equal(data.organizations,1);assert.equal(data.hosts['是'],1);assert.equal(data.attendance['線上'],1);
 assert.deepEqual(data.filters,{profession:'藥事',attendance:'線上',host:'是',from:'',to:''});
});
test('empty valid sheet and filters with no matches return confirmed zero data',()=>{
 for(const {stats} of [harness([headers.slice()]),harness()]){
  const data=stats({profession:'語言治療'});assert.equal(data.ok,true);assert.equal(data.total,0);assert.equal(data.organizations,0);
  assert.equal(data.lastRegistrationAt,null);assert.deepEqual(data.daily,[]);
 }
});
test('invalid date, reversed range and unknown filters reject without reading sheet',()=>{
 for(const params of [{from:'2026-02-30'},{from:'2026-10-09',to:'2026-10-08'},{profession:'unsupported'},
  {attendance:'混合'},{host:'也許'},{to:'2026-1-1'}]){
  const {stats,state}=harness();assert.equal(stats(params).ok,false);assert.equal(state.reads,0);assert.equal(state.writes,0);
 }
});
test('missing or mismatched sheet reports failure without creating or altering data',()=>{
 const data=fixture();data[0][0]='wrong';const original=JSON.stringify(data);
 const {stats,state}=harness(data);assert.equal(stats().ok,false);assert.equal(JSON.stringify(data),original);assert.equal(state.writes,0);
 const missing=harness([], {missing:true});assert.equal(missing.stats().ok,false);assert.equal(missing.state.writes,0);
 const denied=harness(undefined,{accessDenied:true}).stats();assert.equal(denied.ok,false);assert.doesNotMatch(denied.message,/private error/);
});
test('JSON and JSONP payloads contain only aggregates, never names or contact fields',()=>{
 const {get}=harness();const plain=get();assert.equal(plain.type,'json');
 const response=get({callback:'workshopDashboard_test_1'});assert.equal(response.type,'javascript');
 let parsed;vm.runInNewContext(response.body,{workshopDashboard_test_1(data){parsed=data;}});assert.equal(parsed.total,5);
 for(const text of ['測試甲','機構甲','a@example.com','0911111111','id-a','requestId','spreadsheetId']){
  assert.ok(!response.body.includes(text),`Leaked ${text}`);
 }
});
test('malicious or unsupported JSONP callbacks are never reflected as executable code',()=>{
 for(const callback of ['alert(1)','x;alert(1)//','<script>bad</script>','workshopDashboard_'+'a'.repeat(81)]){
  const {get,state}=harness();const result=get({callback});assert.equal(result.type,'json');assert.equal(JSON.parse(result.body).ok,false);
  assert.ok(!result.body.includes(callback));assert.equal(state.reads,0);
 }
});
test('date parser validates real dates, midnight and Sheets Date values in Taipei timezone',()=>{
 const {context}=harness();assert.equal(context.registrationTime_('2026/02/30 12:00:00'),null);
 assert.equal(context.registrationTime_('2026/10/09 24:00:00'),null);
 assert.equal(context.registrationTime_('2026-10-09').timestamp,'2026-10-09 00:00:00');
 assert.equal(context.registrationTime_(new Date('2026-10-08T16:00:00Z')).date,'2026-10-09');
});

test('capacity is global and remains full even when filtered charts show zero registrations',()=>{
 const data=harness().stats({profession:'語言治療'});
 assert.equal(data.total,0);assert.deepEqual(data.capacity,{limit:4,registered:5,remaining:0,full:true});
 const empty=harness([headers.slice()]).stats();assert.deepEqual(empty.capacity,{limit:4,registered:0,remaining:4,full:false});
 const three=harness(fixture().slice(0,4)).stats();assert.deepEqual(three.capacity,{limit:4,registered:3,remaining:1,full:false});
});
test('capacity GET and JSONP contain only counts and match the dashboard capacity',()=>{
 const {context,stats,state}=harness();
 const plain=context.doGet({parameter:{action:'capacity'}});const data=JSON.parse(plain.body);
 assert.equal(data.type,'workshop-capacity');assert.equal(data.ok,true);assert.deepEqual(data.capacity,stats().capacity);
 const output=context.doGet({parameter:{action:'capacity',callback:'workshopCapacity_test_1'}});let received;
 vm.runInNewContext(output.body,{workshopCapacity_test_1(value){received=value;}});assert.equal(received.capacity.remaining,0);
 for(const text of ['測試甲','a@example.com','0911111111','id-a','professions','機構甲'])assert.ok(!output.body.includes(text));
 assert.equal(state.writes,0);
});
test('capacity endpoint rejects unsafe callbacks and reports unavailable data without pretending zero',()=>{
 const {context,state}=harness();const invalid=context.doGet({parameter:{action:'capacity',callback:'alert(1)'}});
 assert.equal(JSON.parse(invalid.body).ok,false);assert.equal(invalid.type,'json');assert.equal(state.reads,0);
 const denied=harness(undefined,{accessDenied:true}).context.doGet({parameter:{action:'capacity'}});
 const data=JSON.parse(denied.body);assert.equal(data.ok,false);assert.equal(data.capacity,undefined);
});
