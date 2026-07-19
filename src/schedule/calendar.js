const MONTHS = ["Январь","Февраль","Март","Апрель","Май","Июнь","Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"];
const DAYS=["Пн","Вт","Ср","Чт","Пт","Сб","Вс"];
export function parseDate(value){if(value instanceof Date&&!Number.isNaN(value.getTime()))return new Date(value);if(typeof value==="number")return new Date(Date.UTC(1899,11,30)+value*86400000);const d=new Date(value);return Number.isNaN(d.getTime())?null:d;}
export function mondayOf(date){const d=new Date(date);const shift=(d.getDay()+6)%7;d.setHours(0,0,0,0);d.setDate(d.getDate()-shift);return d;}
export function buildSemesterCalendar(start,end,workingDays=6){const first=mondayOf(parseDate(start)||new Date());const last=parseDate(end)||first;const rows=[];let week=1;for(let monday=new Date(first);monday<=last;monday.setDate(monday.getDate()+7),week+=1){for(let i=0;i<workingDays;i+=1){const date=new Date(monday);date.setDate(monday.getDate()+i);if(date>last)continue;rows.push({week,day:DAYS[i],date,month:MONTHS[date.getMonth()],dayOfMonth:date.getDate(),working:true});}}return rows;}
export function calendarKey(week,day){return `${Number(week)}|${String(day)}`;}
