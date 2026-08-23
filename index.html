/**
 * Knowledge Speedrun ↔ Google Sheets bridge
 *
 * 1. Create/open your Google Sheet.
 * 2. Extensions → Apps Script.
 * 3. Paste this entire file.
 * 4. Set SHEET_NAME if desired.
 * 5. Deploy → New deployment → Web app.
 *
 * The web app exposes:
 *   GET  ?action=get     -> returns the question bank
 *   POST {action:"replace", questions:[...]} -> replaces question rows
 */
const SHEET_NAME = "KnowledgeBank";

const CSV_ANSWER_COLUMNS = 10;
const CSV_CHOICE_COLUMNS = 25;

const HEADERS = [
  "ID","Category","Subcategory","Question","Type","Difficulty","Correct Answer",
  ...Array.from({length:CSV_ANSWER_COLUMNS},(_,i)=>`Accepted Answer ${i+1}`),
  ...Array.from({length:CSV_CHOICE_COLUMNS},(_,i)=>`Choice ${i+1}`),
  "Pool Count","Displayed Choices","All of the Above","None of the Above","Updated At"
];

function doGet(e) {
  try {
    if ((e && e.parameter && e.parameter.action) !== "get") {
      return json_({ok:true,service:"Knowledge Speedrun Sheet Bridge"});
    }
    return json_({ok:true,questions:readQuestions_()});
  } catch (err) {
    return json_({ok:false,error:String(err && err.message || err)});
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error("No POST data.");
    const body = JSON.parse(e.postData.contents);
    if (body.action !== "replace" || !Array.isArray(body.questions)) {
      throw new Error("Expected {action:'replace', questions:[...]}.");
    }
    writeQuestions_(body.questions);
    return json_({ok:true,count:body.questions.length});
  } catch (err) {
    return json_({ok:false,error:String(err && err.message || err)});
  }
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  return sh;
}

function readQuestions_() {
  const sh = sheet_();
  const lastRow = sh.getLastRow();
  const lastCol = Math.max(sh.getLastColumn(), HEADERS.length);
  if (lastRow < 2) return [];
  const values = sh.getRange(1,1,lastRow,lastCol).getDisplayValues();
  const header = values[0];
  const col = {};
  header.forEach((h,i)=>col[h]=i);

  return values.slice(1).filter(r => r.some(x=>String(x).trim()!=="")).map(r=>{
    const get = name => col[name] == null ? "" : String(r[col[name]] ?? "").trim();
    const type = get("Type").toLowerCase() || "typed";
    const q = {
      id:get("ID"),
      category:get("Category"),
      subcategory:get("Subcategory"),
      question:get("Question"),
      type:type === "multiple choice" ? "mc" : type,
      difficulty:get("Difficulty") || "normal",
      updatedAt:Number(get("Updated At")) || Date.now()
    };
    if(q.type === "mc"){
      q.choices = [];
      for(let i=1;i<=CSV_CHOICE_COLUMNS;i++){
        const v=get("Choice "+i); if(v) q.choices.push(v);
      }
      q.correct=get("Correct Answer");
      q.poolCount=Math.min(25,Math.max(2,parseInt(get("Pool Count"),10)||q.choices.length||4));
      q.displayCount=Math.min(25,Math.max(2,parseInt(get("Displayed Choices"),10)||4));
      q.allAbove=/^(true|yes|1)$/i.test(get("All of the Above"));
      q.noneAbove=/^(true|yes|1)$/i.test(get("None of the Above"));
    }else{
      q.answers=[];
      for(let i=1;i<=CSV_ANSWER_COLUMNS;i++){
        const v=get("Accepted Answer "+i); if(v) q.answers.push(v);
      }
      if(!q.answers.length && get("Correct Answer")) q.answers=[get("Correct Answer")];
    }
    return q;
  });
}

function writeQuestions_(questions) {
  const sh=sheet_();
  sh.clearContents();
  sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);

  if(!questions.length){
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1,HEADERS.length);
    return;
  }

  const rows=questions.map(q=>{
    const answers=q.type==="typed" ? (q.answers||[]) : [];
    const choices=q.type==="mc" ? (q.choices||[]) : [];
    return [
      q.id||Utilities.getUuid(),q.category||"",q.subcategory||"",q.question||"",
      q.type||"typed",q.difficulty||"normal",q.correct||answers[0]||"",
      ...Array.from({length:CSV_ANSWER_COLUMNS},(_,i)=>answers[i]||""),
      ...Array.from({length:CSV_CHOICE_COLUMNS},(_,i)=>choices[i]||""),
      q.poolCount||choices.length||0,q.displayCount||4,
      q.allAbove?"true":"false",q.noneAbove?"true":"false",
      Number(q.updatedAt)||Date.now()
    ];
  });
  sh.getRange(2,1,rows.length,HEADERS.length).setValues(rows);
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,HEADERS.length).setFontWeight("bold");
  sh.autoResizeColumns(1,HEADERS.length);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
