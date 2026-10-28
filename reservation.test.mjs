import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {readFile} from 'node:fs/promises';
const source=await readFile(new URL('Code.gs',import.meta.url),'utf8');

test('11-column reservation layout reads and updates without overwriting status or contact checkbox',()=>{
 const c=vm.createContext({Utilities:{formatDate:()=> '2026-10-28 12:00:00'}});vm.runInContext(source,c);
 const row=['Test Member','2026/11/1','14:00','(DC指定省略)','コンパニオンアプリ','','old','要対応',true,'随時受付','保持メモ / 1時間予約'];
 const sheet={getLastRow:()=>2,getLastColumn:()=>11,getRange:(r,col,height,width)=>({
  getDisplayValues:()=>[row.slice(col-1,col-1+width).map(String)],
  setValues:values=>{assert.equal(values[0].length,width);values[0].forEach((v,i)=>row[col-1+i]=v);},
  setValue:value=>{row[col-1]=value;},sort:()=>{}
 })};
 let read=c.readReservations_(sheet)[0];assert.equal(read.supportMethod,'companion');assert.equal(read.mode,'flexible');assert.equal(read.oneHourBooking,true);
 c.upsertReservation_(sheet,{row:2},{name:'Test Member',date:'2026-11-02',time:'15:00',supportMethod:'game',mode:'normal',oneHourBooking:false,remarks:'保持メモ'});
 assert.equal(row[4],'ゲーム内');assert.equal(row[6],'2026-10-28 12:00:00');assert.equal(row[7],'要対応');assert.equal(row[8],true);assert.equal(row[9],'通常予約');assert.equal(row[10],'保持メモ');
 read=c.readReservations_(sheet)[0];assert.equal(read.date,'2026-11-02');assert.equal(read.oneHourBooking,false);assert.equal('meetingWorld' in read,false);
});
