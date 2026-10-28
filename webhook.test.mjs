import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {readFile} from 'node:fs/promises';
import {randomUUID} from 'node:crypto';

async function fixture(){
  const rows=[],sent=[],properties=new Map(),triggers=[];let locked=false,status=201;
  const sheet={appendRow:row=>rows.push([...row]),setFrozenRows:()=>{},getDataRange:()=>({getValues:()=>rows.map(row=>[...row])}),getRange:(row,column)=>({setValues:values=>values.forEach((value,index)=>value.forEach((cell,offset)=>rows[row-1+index][column-1+offset]=cell))})};
  let exists=false;
  const context=vm.createContext({console,Date,PropertiesService:{getScriptProperties:()=>({getProperty:name=>properties.get(name)??null})},Utilities:{getUuid:randomUUID},
    LockService:{getScriptLock:()=>({tryLock:()=>{assert.equal(locked,false);locked=true;return true;},releaseLock:()=>{locked=false;}})},
    SpreadsheetApp:{flush:()=>{}},UrlFetchApp:{fetch:(url,options)=>{assert.equal(locked,false,'delivery must not hold the reservation lock');sent.push({url,options});return {getResponseCode:()=>status,getAllHeaders:()=>({'Retry-After':'60'})};}},
    ScriptApp:{getProjectTriggers:()=>triggers,newTrigger:handler=>({timeBased(){return this;},everyMinutes(value){assert.equal(value,1);return this;},create(){triggers.push({getHandlerFunction:()=>handler});}}),deleteTrigger:item=>triggers.splice(triggers.indexOf(item),1)}});
  for(const file of ['Code.gs','BoardWebhook.gs'])vm.runInContext(await readFile(new URL(file,import.meta.url),'utf8'),context,{filename:file});
  context.getReservationSpreadsheet_=()=>({getSheetByName:()=>exists?sheet:null,insertSheet:()=>{exists=true;return sheet;}});
  properties.set('GCAL_BOARD_WEBHOOK_URL','https://notify.invalid/hooks/board/'+randomUUID()+'/'+'x'.repeat(43));
  return {context,rows,sent,properties,triggers,setStatus:value=>status=value};
}

test('booking webhook queue keeps reservation events distinct and delivers after releasing locks',async()=>{
  const f=await fixture(),c=f.context;
  c.installBoardWebhookDelivery();c.installBoardWebhookDelivery();assert.equal(f.triggers.length,1);
  const reservation={name:'Player_* One',date:'2026-10-18',time:'20:00',meetingDc:'Gaia',meetingWorld:'Bahamut',supportMethod:'game',oneHourBooking:true};
  c.notifyDiscordReservationCreated_(reservation);c.notifyDiscordReservationChanged_({...reservation,time:'19:00'},reservation);c.notifyDiscordReservationCancelled_(reservation);
  c.notifyDiscordAutoCancelled_([reservation,reservation],[]);
  assert.equal(f.rows.length,6);assert.equal(new Set(f.rows.slice(1).map(row=>row[0])).size,5);assert.equal(f.sent.length,0,'booking functions queue without sending board HTTP');
  c.retryBoardWebhookNotifications();assert.equal(f.sent.length,5);assert.ok(f.rows.slice(1).every(row=>row[3]==='送信済み'));
  for(const sent of f.sent){const data=JSON.parse(sent.options.payload);assert.deepEqual(Object.keys(data),['content']);assert.match(data.content,/Player\\_\\\* One/);assert.ok(sent.options.headers['Idempotency-Key']);assert.equal(sent.options.followRedirects,false);}
  c.retryBoardWebhookNotifications();assert.equal(f.sent.length,5);
});

test('booking webhook retry preserves notification ID and separates input errors',async()=>{
  const f=await fixture(),c=f.context;c.installBoardWebhookDelivery();
  c.queueBoardReservation_('予約',{name:'Test',supportMethod:'game'});f.setStatus(503);c.retryBoardWebhookNotifications();
  assert.equal(f.rows[1][3],'待機');assert.equal(f.rows[1][4],1);const first=f.sent[0].options.headers['Idempotency-Key'];
  f.rows[1][5]=new Date(0);f.setStatus(201);c.retryBoardWebhookNotifications();assert.equal(f.rows[1][3],'送信済み');assert.equal(f.sent[1].options.headers['Idempotency-Key'],first);
  c.queueBoardReservation_('予約',{name:'Test',supportMethod:'game'});f.setStatus(400);c.retryBoardWebhookNotifications();assert.equal(f.rows[2][3],'要確認');const count=f.sent.length;c.retryBoardWebhookNotifications();assert.equal(f.sent.length,count);
});

test('booking webhook remains inactive without a configured secret URL',async()=>{
  const f=await fixture();f.properties.delete('GCAL_BOARD_WEBHOOK_URL');f.context.notifyDiscordReservationCreated_({name:'Test'});assert.equal(f.rows.length,0);assert.equal(f.sent.length,0);assert.throws(()=>f.context.installBoardWebhookDelivery());
});

test('Discord and board share concise fields, time ranges and changes without world or empty contacts',async()=>{
 const f=await fixture(),c=f.context;f.properties.set('DISCORD_WEBHOOK_URL','https://discord.invalid/test');
 const before={name:'Test Member',date:'2026-11-01',time:'14:00',supportMethod:'game',mode:'normal',oneHourBooking:false};
 const after={...before,time:'15:00',supportMethod:'companion',oneHourBooking:true,meetingWorld:'legacy ignored'};
 c.notifyDiscordReservationCreated_(after);c.notifyDiscordReservationChanged_(before,after);c.notifyDiscordReservationCancelled_(after);
 const embeds=f.sent.map(item=>JSON.parse(item.options.payload).embeds[0]);
 assert.equal(embeds.length,3);
 const titles=['🟢 新規予約','🔄 予約変更','❌ 予約取り消し'];
 assert.deepEqual(embeds.map(embed=>embed.title),titles);
 titles.forEach((title,index)=>assert.ok(f.rows[index+1][2].startsWith('## '+title+'\n\n')));
 for(const embed of embeds){
  const names=embed.fields.map(field=>field.name);assert.equal(names.includes('集合ワールド'),false);assert.equal(names.includes('集合DC'),false);assert.equal(names.includes('連絡先'),false);
  assert.equal(embed.fields.find(field=>field.name==='予約日時').value,'2026-11-01 15:00～16:00');
  assert.equal(embed.fields.find(field=>field.name==='予約時間').value,'60分');
 }
 const changed=embeds[1].fields;
 assert.equal(changed.find(field=>field.name==='変更前日時').value,'2026-11-01 14:00～14:30');assert.match(changed.find(field=>field.name==='変更内容').value,/ゲーム内 → コンパニオンアプリ/);
 const board=c.boardReservationText_('予約変更',after,before);assert.doesNotMatch(board,/ワールド|連絡先：なし/);assert.equal((c.boardReservationText_('新規予約',after).match(/予約時間：/g)||[]).length,1);assert.match(board,/受付方式：\*\* 通常予約/);
 assert.equal(c.hasReservationChanged_(before,{...before,oneHourBooking:true}),true);assert.equal(c.hasReservationChanged_(before,{...before}),false);
 assert.equal(c.reservationDateTime_({...before,time:'23:30'}),'2026-11-01 23:30～24:00');
 const historical=c.reservationFields_({...before,supportMethod:'x',contact:'@old'});assert.equal(historical.find(field=>field.name==='連絡先').value,'@old');
});

test('automatic cancellations include reasons on both destinations and split Discord embeds by 10',async()=>{
 const f=await fixture(),c=f.context;f.properties.set('DISCORD_WEBHOOK_URL','https://discord.invalid/test');
 const reservation={name:'Test Member',date:'2026-11-01',time:'14:00',supportMethod:'game'};
 const period={start:'2026-11-01T13:00',end:'2026-11-01T16:00',startMs:c.localStampToPseudoMs_('2026-11-01 13:00'),endMs:c.localStampToPseudoMs_('2026-11-01 16:00'),blockGame:true,message:'メンテナンス'};
 c.notifyDiscordAutoCancelled_(Array.from({length:21},()=>({...reservation})),[period]);
 const batches=f.sent.map(item=>JSON.parse(item.options.payload).embeds);assert.deepEqual(batches.map(batch=>batch.length),[10,10,1]);
 assert.ok(batches.flat().every(embed=>embed.fields.some(field=>field.name==='取消理由' && field.value==='ゲーム内停止')));
 assert.equal(f.rows.length,22);assert.match(f.rows[1][2],/取消理由：\*\* ゲーム内停止/);assert.match(f.rows[1][2],/該当時間：/);assert.doesNotMatch(f.rows[1][2],/ワールド/);
});


test('board Markdown decorates labels while escaping literal member text and multiline reasons',async()=>{
 const f=await fixture(),c=f.context;
 const text=c.boardReservationText_('予約変更',{name:'Player_* [One]',date:'2026-11-01',time:'14:00',supportMethod:'game'},null,[{name:'取消理由',value:'点検\n**入力内容**'}]);
 assert.ok(text.startsWith('## 予約変更\n\n'));
 assert.match(text,/- \*\*キャラクター名：\*\* Player\\_\\\* \\\[One\\\]/);
 assert.match(text,/- \*\*予約日時：\*\* 2026\\-11\\-01 14:00～14:30/);
 assert.match(text,/- \*\*取消理由：\*\* 点検\n  - \\\*\\\*入力内容\\\*\\\*/);
});
