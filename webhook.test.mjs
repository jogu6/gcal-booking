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
