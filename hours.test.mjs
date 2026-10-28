import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {readFile} from 'node:fs/promises';
const source=await readFile(new URL('Code.gs',import.meta.url),'utf8');
const keys=['日','月','火','水','木','金','土','祝日'];
const defaults=()=>keys.map(key=>[key,'09:00','23:00']);
function context(){
 const c=vm.createContext({console,Utilities:{formatDate:(date,tz,pattern)=>{
  const p=new Intl.DateTimeFormat('en-CA',{timeZone:tz,year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',hourCycle:'h23'}).formatToParts(date);
  const v=Object.fromEntries(p.map(x=>[x.type,x.value]));
  return pattern==='HH:mm'?`${v.hour}:${v.minute}`:`${v.year}-${v.month}-${v.day}`;
 }}});vm.runInContext(source,c);return c;
}
function book(rows=defaults()) {return {getSheetByName:name=>rows===null?null:({getLastRow:()=>rows.length+1,getMaxColumns:()=>Math.max(3,...rows.map(row=>row.length)),getRange:()=>({getDisplayValues:()=>rows})})};}

test('weekday settings, holiday override, missing-tab compatibility and days off',()=>{
 const c=context(),rows=defaults();rows[1]=['月','18:00','22:00'];rows[7]=['祝日','10:00','20:00'];rows[2]=['火','',''];
 const h=c.getDailyBookingHours_(book(rows),'2026-11-02','2026-11-04',['2026-11-03']);
 assert.equal(h['2026-11-02'].intervals[0].startMinute,1080);assert.equal(h['2026-11-02'].intervals[0].endMinute,1320);
 assert.equal(h['2026-11-03'].category,'祝日');assert.equal(h['2026-11-03'].intervals[0].startMinute,600);
 assert.equal(c.getDailyBookingHours_(book(rows),'2026-11-03','2026-11-03',[])['2026-11-03'].intervals.length,0);
 assert.equal(c.getWeeklyBookingHours_(book(null))['日'].intervals[0].endMinute,1380);
 assert.deepEqual(Object.keys(c.getDailyBookingHours_(book(),'2026-11-04','2026-11-02',[])),[]);
});

test('invalid, partial, duplicate and missing schedule rows fail closed',()=>{
 const c=context();
 for(const time of ['09:65','25:00','abc','24:00']){const rows=defaults();rows[0][1]=time;assert.throws(()=>c.getWeeklyBookingHours_(book(rows)));}
 for(const [start,end]of [['23:00','09:00'],['09:00','09:00'],['','23:00'],['09:00','']]){const rows=defaults();rows[0]=['日',start,end];assert.throws(()=>c.getWeeklyBookingHours_(book(rows)));}
 assert.throws(()=>c.getWeeklyBookingHours_(book(defaults().slice(1))));
 assert.throws(()=>c.getWeeklyBookingHours_(book([...defaults(),['日','09:00','23:00']])));
 assert.throws(()=>c.getWeeklyBookingHours_(book([['日曜','09:00','23:00'],...defaults().slice(1)])));
});

test('30/60 minute boundaries and midnight closing are enforced',()=>{
 const c=context(),h={intervals:[{startMinute:1080,endMinute:1320}],stepMinutes:30};
 assert.equal(c.isWithinBookingHours_(h,'17:30',30),false);
 assert.equal(c.isWithinBookingHours_(h,'18:00',30),true);
 assert.equal(c.isWithinBookingHours_(h,'21:30',30),true);
 assert.equal(c.isWithinBookingHours_(h,'21:30',60),false);
 assert.equal(c.isWithinBookingHours_(h,'21:00',60),true);
 assert.equal(c.isWithinBookingHours_(h,'22:00',30),false);
 assert.equal(c.isWithinBookingHours_(h,'18:15',30),false);
 assert.equal(c.isWithinBookingHours_({intervals:[{startMinute:0,endMinute:1440}],stepMinutes:30},'23:30',30),true);
 assert.equal(c.isWithinBookingHours_({intervals:[{startMinute:0,endMinute:1440}],stepMinutes:30},'23:00',60),true);
 assert.equal(c.isWithinBookingHours_({intervals:[],stepMinutes:30},'00:00',30),false);
});

function bookingFixture(oneHour, customRows){
 const c=context(),rows=defaults();rows.forEach(row=>{row[1]='18:00';row[2]='22:00';});
 const ss=book(customRows || rows);c.getReservationSpreadsheet_=()=>ss;c.getBookingWindow_=()=>({start:'2099-01-01',end:'2099-01-07'});
 c.getHolidayDates_=()=>[];c.getSheetFrom_=()=>({});c.getMaintenancePeriodsFromSheet_=()=>[];c.getSpecialReceptionPeriodsFromSheet_=()=>[];
 c.getMemberSheet_=()=>({});c.findMember_=()=>({name:'Test Member',status:'ご予約待ち/ご予約中',oneHourBooking:oneHour});
 c.readReservations_=()=>[];let saved=null;c.upsertReservation_=(sheet,current,data)=>{saved=data;};c.shouldNotifyAllReservationEvents_=()=>false;
 c.LockService={getScriptLock:()=>({waitLock:()=>{},releaseLock:()=>{}})};
 return {c,input:time=>({name:'Test Member',date:'2099-01-01',time,supportMethod:'game'}),saved:()=>saved};
}

test('actual booking path accepts the last valid slots and rejects out-of-hours or retired X without writes',()=>{
 let f=bookingFixture(false);assert.equal(f.c.book_(f.input('21:30')).ok,true);assert.ok(f.saved());
 f=bookingFixture(true);assert.equal(f.c.book_(f.input('21:00')).durationMinutes,60);
 f=bookingFixture(true);assert.equal(f.c.book_(f.input('21:30')).code,'ONE_HOUR_OUTSIDE_HOURS');assert.equal(f.saved(),null);
 f=bookingFixture(false);assert.throws(()=>f.c.book_(f.input('17:30')),/予約可能時間外/);assert.equal(f.saved(),null);
 f=bookingFixture(false);assert.throws(()=>f.c.book_({...f.input('20:00'),supportMethod:'x'}),/受付は終了/);assert.equal(f.saved(),null);
});

test('holiday master returns ISO text and Date cells with inclusive boundaries',()=>{
 const c=context();
 const values=['2026-01-01','2026-11-03','2026-11-23','2027-01-01'];
 c.SpreadsheetApp={openById:()=>({getSheetByName:()=>({getLastRow:()=>values.length+1,getRange:(row,col,height)=>({
   getValue:()=>values[row-2],
   createTextFinder:pattern=>({useRegularExpression(){return this;},matchEntireCell(){return this;},findNext(){const i=values.findIndex(v=>typeof v==='string'&&new RegExp(pattern).test(v));return i<0?null:{getRow:()=>i+2};}})
 })})})};
 assert.equal(JSON.stringify(c.getHolidayDates_('2026-11-01','2026-11-23')),JSON.stringify(['2026-11-03','2026-11-23']));
 assert.equal(JSON.stringify(c.getHolidayDates_('2026-11-03','2026-11-03')),JSON.stringify(['2026-11-03']));
 values[1]=vm.runInContext("new Date('2026-11-03T00:00:00+09:00')",c);
 assert.equal(JSON.stringify(c.getHolidayDates_('2026-11-01','2026-11-03')),JSON.stringify(['2026-11-03']));
});

test('historical X rows retain their method and contact and remain cancelable',()=>{
 const c=context(),rows=[['Test Member','2099/1/1','20:00','','','X','@test','','','','通常予約','']];
 const sheet={getLastRow:()=>2,getRange:()=>({getDisplayValues:()=>rows}),deleteRow:row=>{assert.equal(row,2);rows.length=0;}};
 const reservation=c.readReservations_(sheet)[0];assert.equal(reservation.supportMethod,'x');assert.equal(reservation.contact,'@test');
 c.getReservationSheet_=()=>sheet;c.shouldNotifyAllReservationEvents_=()=>false;c.LockService={getScriptLock:()=>({waitLock:()=>{},releaseLock:()=>{}})};
 assert.equal(c.cancel_('Test Member').ok,true);assert.equal(rows.length,0);
});


test('up to three sorted intervals, gaps, adjacent boundaries and custom start steps',()=>{
 const c=context(),rows=defaults();rows[1]=['月','19:00','22:00','09:00','12:00','14:00','16:00','15'];
 const h=c.getDailyBookingHours_(book(rows),'2026-11-02','2026-11-02',[])['2026-11-02'];
 assert.equal(h.intervals.length,3);assert.equal(h.intervals[0].startMinute,540);assert.equal(h.stepMinutes,15);
 assert.equal(c.isWithinBookingHours_(h,'09:15',30),true);
 assert.equal(c.isWithinBookingHours_(h,'11:45',30),false);
 assert.equal(c.isWithinBookingHours_(h,'12:30',30),false);
 assert.equal(c.isWithinBookingHours_(h,'15:30',60),false);
 assert.equal(c.isWithinBookingHours_(h,'21:00',60),true);
 assert.equal(c.candidateBookingTimes_(h).includes('09:15'),true);
 assert.equal(c.candidateBookingTimes_(h).includes('12:00'),false);
 rows[1]=['月','09:00','10:00','10:00','11:00','','','15'];
 const adjacent=c.getWeeklyBookingHours_(book(rows))['月'];
 assert.equal(c.isWithinBookingHours_(adjacent,'09:30',60),false);
 rows[1]=['月','09:10','10:40','','','','','30'];
 assert.equal(c.candidateBookingTimes_(c.getWeeklyBookingHours_(book(rows))['月']).join(','),'09:10,09:40,10:10');
 rows[1][7]='60';assert.equal(c.candidateBookingTimes_(c.getWeeklyBookingHours_(book(rows))['月']).join(','),'09:10,10:10');
});

test('overlapping/partial intervals and unsupported steps are rejected',()=>{
 const c=context();
 for(const row of [ ['月','09:00','12:00','11:30','13:00'],['月','09:00','12:00','14:00',''],['月','09:00','12:00','','','','','10'] ]){
  const rows=defaults();rows[1]=row;assert.throws(()=>c.getWeeklyBookingHours_(book(rows)));
 }
});

test('15-minute candidate grid blocks all starts overlapping 30/60-minute reservations',()=>{
 const c=context(),window={start:'2099-01-01',end:'2099-01-01'},hours={'2099-01-01':{intervals:[{startMinute:540,endMinute:720}],stepMinutes:15}};
 const r={date:'2099-01-01',time:'10:00',mode:'normal',oneHourBooking:false};
 const booked=c.getBookedSlotsFromReservations_([r],window,[],hours);
 assert.equal(booked.join(','),'2099-01-01 09:45,2099-01-01 10:00,2099-01-01 10:15');
 assert.equal(c.reservationOverlaps_(r,r.date,'09:30',30),false);
 assert.equal(c.reservationOverlaps_(r,r.date,'10:30',30),false);
 r.oneHourBooking=true;
 assert.equal(c.getBookedSlotsFromReservations_([r],window,[],hours).length,5);
 const rows=defaults().map(row=>[row[0],'09:00','12:00','','','','','15']);
 const f=bookingFixture(false,rows);f.c.readReservations_=()=>[{...r,row:2,normalizedName:'other member'}];
 assert.equal(f.c.book_(f.input('09:45')).code,'SLOT_TAKEN');assert.equal(f.saved(),null);
 assert.equal(f.c.book_(f.input('11:00')).ok,true);
});

test('one-hour booking cannot span a gap; flexible reception keeps its sharing policy',()=>{
 const rows=defaults().map(row=>[row[0],'18:00','19:00','19:30','22:00','','','15']);
 let f=bookingFixture(true,rows);assert.equal(f.c.book_(f.input('18:30')).code,'ONE_HOUR_OUTSIDE_HOURS');assert.equal(f.saved(),null);
 f=bookingFixture(false,rows);f.c.getSpecialReceptionStateForSlot_=()=>({state:'flexible'});
 f.c.readReservations_=()=>[{date:'2099-01-01',time:'20:00',mode:'flexible',row:2,normalizedName:'other member'}];
 assert.equal(f.c.book_(f.input('20:15')).ok,true);
});


test('flexible reservations cannot appear free across a transition to normal booking',()=>{
 const c=context(),window={start:'2099-01-01',end:'2099-01-01'},hours={'2099-01-01':{intervals:[{startMinute:540,endMinute:720}],stepMinutes:15}};
 c.getSpecialReceptionStateForSlot_=(date,time)=>time<'10:00'?{state:'flexible'}:null;
 const r={date:'2099-01-01',time:'09:45',mode:'flexible'};
 const blocked=c.getBookedSlotsFromReservations_([r],window,[],hours);
 assert.equal(blocked.join(','),'2099-01-01 10:00');
 const counts=c.getFlexibleSlotCounts_([r],window,[],hours);
 assert.equal(counts['2099-01-01 09:45'],1);assert.equal(counts['2099-01-01 10:00'],undefined);
});
