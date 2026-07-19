export function normalizeId(value){return String(value??"").trim();}
export function lessonKey(lesson){return [lesson.sourceId,lesson.group,lesson.week,lesson.day??lesson.dayOfWeek,lesson.pair??lesson.pairNumber,lesson.cellAddress].map(normalizeId).join("|");}
export function slotKey(week,day,pair){return [Number(week),String(day),Number(pair)].join("|");}
export function streamKey(lesson){return [lesson.week,lesson.day??lesson.dayOfWeek,lesson.pair??lesson.pairNumber,lesson.subject,lesson.lessonType,lesson.room].map(normalizeId).join("|");}
export function assignmentId(lessonId){return `A:${normalizeId(lessonId)}`;}
export function asBoolean(value){if(typeof value==="boolean")return value;return /^(?:1|true|yes|да|истина)$/i.test(String(value??"").trim());}
