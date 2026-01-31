/** AI Grader - criteria based (corrected)
 * Paste into Apps Script editor bound to your Master spreadsheet.
 */

/** --- CONFIG --- */
const SLEEP_BETWEEN_CALLS_MS = 1000;
const HF_API_URL_PREFIX = "https://router.huggingface.co/hf-inference/models/";
const GRADE_SHEET_NAME = "AI_GRADES";
const ANSWER_KEY_SHEET = "ANSWER_KEY";
const MASTER_CONFIG_SHEET = "CONFIG";

/** Read CONFIG from the Master spreadsheet (opened by ID saved in properties) */
function readConfig() {
  const masterId = PropertiesService.getScriptProperties().getProperty('MASTER_SPREADSHEET_ID');
  if (!masterId) throw new Error("MASTER_SPREADSHEET_ID not set. Run installTriggers() from the Master spreadsheet first.");
  const ss = SpreadsheetApp.openById(masterId);
  const sheet = ss.getSheetByName(MASTER_CONFIG_SHEET);
  if (!sheet) throw new Error("CONFIG sheet not found in Master spreadsheet.");
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];
  const headers = rows.shift();
  const idx = (h) => headers.indexOf(h);
  return rows.map(r => ({
    responseSpreadsheetId: r[idx("response_spreadsheet_id")],
    responseSheetName: r[idx("response_sheet_name")],
    emailColumnHeader: r[idx("email_column_header")],
    examName: r[idx("exam_name")],
    hfModel: r[idx("hf_model")],
    nameColumnHeader: r[idx("name_column_header")]
  })).filter(x => x.responseSpreadsheetId);
}

/** Install triggers for all configured response spreadsheets.
 * Also stores the Master spreadsheet id in Script Properties so triggers can later read CONFIG.
 */
function installTriggers() {
  const masterSs = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('MASTER_SPREADSHEET_ID', masterSs.getId());
  const configs = readConfig(); // now reads the master by id we just stored

  // remove existing triggers for handleFormSubmit to avoid duplicates
  const allTriggers = ScriptApp.getProjectTriggers();
  allTriggers.forEach(t => {
    if (t.getHandlerFunction() === "handleFormSubmit") {
      ScriptApp.deleteTrigger(t);
    }
  });

  configs.forEach(cfg => {
    try {
      // open spreadsheet object for trigger builder
      const targetSs = SpreadsheetApp.openById(cfg.responseSpreadsheetId);
      ScriptApp.newTrigger("handleFormSubmit")
        .forSpreadsheet(targetSs)
        .onFormSubmit()
        .create();
      Logger.log("Trigger created for spreadsheet: " + cfg.responseSpreadsheetId);
    } catch (err) {
      Logger.log("Failed to create trigger for " + cfg.responseSpreadsheetId + ": " + err);
    }
  });
}

/** Main handler invoked by form submit trigger.
 * e.source is the Response spreadsheet where the form wrote the row.
 */

function handleFormSubmit(e) {
  try {
    const configs = readConfig();
    const ssId = e.source.getId();
    const cfg = configs.find(c => c.responseSpreadsheetId == ssId);
    if (!cfg) {
      Logger.log("No config for spreadsheet id: " + ssId);
      return;
    }

    const namedValues = e.namedValues || {};
    const emailHeader = cfg.emailColumnHeader || "Email Address";
    const studentEmailArr = namedValues[emailHeader];
    const studentEmail = studentEmailArr && studentEmailArr.length ? studentEmailArr[0] : "";

    // Attempt to find student name
    let studentName = "Student";
    const nameHeader = cfg.nameColumnHeader;
    if (nameHeader && namedValues[nameHeader]) {
      studentName = namedValues[nameHeader][0];
    } else {
      const nameKeys = Object.keys(namedValues).filter(k => /name/i.test(k));
      if (nameKeys.length > 0) {
        studentName = namedValues[nameKeys[0]][0];
      }
    }

    const respSS = SpreadsheetApp.openById(ssId);

    /* ================= ANSWER_KEY ================= */
    const answerSheet = respSS.getSheetByName(ANSWER_KEY_SHEET);
    if (!answerSheet) {
      Logger.log("ANSWER_KEY sheet not found");
      return;
    }

    const akData = answerSheet.getDataRange().getValues();
    if (akData.length < 2) return;

    const akHeader = akData.shift();
    const qIndex = akHeader.indexOf("question_header");
    const idIndex = akHeader.indexOf("question_id");
    const criteriaIndex = akHeader.indexOf("criteria_json");

    /* ================= AI_GRADES ================= */
    let gradeSheet = respSS.getSheetByName(GRADE_SHEET_NAME);
    if (!gradeSheet) {
      gradeSheet = respSS.insertSheet(GRADE_SHEET_NAME);
      gradeSheet.appendRow([
        "timestamp",
        "student_email",
        "exam_name",
        "question_id",
        "question_header",
        "score",
        "feedback",
        "criteria_results",
        "improvement",
        "raw_model_output",
        "status"
      ]);
    }

    /* =========================================================
       PART A — BUILD HEADER INDEX MAP
    ========================================================== */
    const responseSheetName = cfg.responseSheetName || "Form Responses 1";
    const responseSheet = respSS.getSheetByName(responseSheetName) || respSS.getSheets()[0];
    const responseHeaders = responseSheet.getRange(1, 1, 1, responseSheet.getLastColumn()).getValues()[0];

    function normalizeHeader(h) {
      return h ? h.toString().replace(/\s+/g, " ").trim().toLowerCase() : "";
    }

    const headerIndexMap = {};
    responseHeaders.forEach((h, i) => {
      headerIndexMap[normalizeHeader(h)] = i;
    });

    const submittedValues = e.values || [];

    /* ================= PREPARE BATCH DATA ================= */
    const itemsToGrade = [];
    for (let r = 0; r < akData.length; r++) {
      const row = akData[r];
      const qHeader = row[qIndex];
      const qId = row[idIndex] || "Q" + (r + 1);
      const criteriaJson = row[criteriaIndex];

      if (!qHeader || !criteriaJson) continue;

      let studentAnswer = "";
      const normQ = normalizeHeader(qHeader);

      if (headerIndexMap.hasOwnProperty(normQ)) {
        studentAnswer = submittedValues[headerIndexMap[normQ]] || "";
      } else if (namedValues[qHeader]) {
        studentAnswer = namedValues[qHeader][0] || "";
      }

      itemsToGrade.push({
        id: qId,
        header: qHeader,
        answer: studentAnswer.trim(),
        criteria: criteriaJson
      });
    }

    if (itemsToGrade.length === 0) return;

    /* ================= BATCH AI GRADING (CHUNKED) ================= */
    const CHUNK_SIZE = 5;
    const batchResults = [];
    let combinedModelOutput = "";

    for (let i = 0; i < itemsToGrade.length; i += CHUNK_SIZE) {
      const chunk = itemsToGrade.slice(i, i + CHUNK_SIZE);
      Logger.log(`Grading chunk ${Math.floor(i/CHUNK_SIZE) + 1} (${chunk.length} questions)...`);
      
      const prompt = buildBatchGradingPrompt(cfg.examName, chunk);
      let chunkOutput = "";
      
      try {
        chunkOutput = callGroqApi(prompt);
        combinedModelOutput += `--- CHUNK ${Math.floor(i/CHUNK_SIZE) + 1} ---\n${chunkOutput}\n\n`;
        
        const parsedChunk = tryParseBatchJson(chunkOutput, chunk.length);
        Logger.log(`Parsed ${parsedChunk.length} results from chunk.`);
        batchResults.push(...parsedChunk);
      } catch (err) {
        Logger.log(`Chunk grading error: ${err}`);
      }
    }

    /* ================= PROCESS RESULTS ================= */
    const rowsToAppend = [];
    const resultsForEmail = [];
    const timestamp = new Date();

    itemsToGrade.forEach((item, index) => {
      // Find matching result by ID
      let result = batchResults.find(r => r.question_id === item.id);
      
      if (!result) {
        result = {
          total_score: 0,
          feedback: item.answer ? "AI could not evaluate this response." : "No answer was provided.",
          criteria_results: [],
          improvement: item.answer ? "Answer all required criteria clearly." : "Please answer the question."
        };
      }

      rowsToAppend.push([
        timestamp,
        studentEmail,
        cfg.examName,
        item.id,
        item.header,
        result.total_score,
        result.feedback,
        JSON.stringify(result.criteria_results),
        result.improvement,
        index === 0 ? combinedModelOutput : "See first row for full batch output",
        "graded_by_ai"
      ]);

      resultsForEmail.push({
        qId: item.id,
        qHeader: item.header,
        score: result.total_score,
        feedback: result.feedback,
        criteriaResults: result.criteria_results,
        improvement: result.improvement
      });
    });

    // Batch write to sheet
    if (rowsToAppend.length > 0) {
      gradeSheet.getRange(gradeSheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
    }

    /* ================= EMAIL ================= */
    if (studentEmail) {
      const htmlBody = buildEmailHtml(cfg.examName, studentName, resultsForEmail);
      MailApp.sendEmail({
        to: studentEmail,
        subject: `Exam Result: ${cfg.examName}`,
        htmlBody: htmlBody
      });
    }

  } catch (err) {
    Logger.log("handleFormSubmit error: " + err);
  }
}


/** Build a prompt for batch grading multiple questions at once. */
function buildBatchGradingPrompt(examName, items) {
  let questionsBlock = items.map((item, idx) => {
    return `--- QUESTION ${item.id} ---
Header: ${item.header}
Student Answer: ${item.answer || "[No answer provided]"}
Criteria (JSON): ${item.criteria}`;
  }).join("\n\n");

  return `
You are an expert exam grader for the exam: "${examName}".
Grade the following ${items.length} questions.

${questionsBlock}

Rules:
1. Grade EACH question strictly according to its unique criteria.
2. Award full marks for a criterion ONLY if clearly mentioned.
3. If partially mentioned, award half marks. If not mentioned, award 0.
4. Use Ethiopian context where relevant.
5. Be concise.

Return ONLY a valid JSON object with a "results" array containing one entry for each question:
{
  "results": [
    {
      "question_id": "Q1",
      "total_score": 16,
      "feedback": "...",
      "improvement": "...",
      "criteria_results": [{"id": "website_type", "awarded": 4, "reason": "..."}]
    },
    ...
  ]
}
`;
}

/** Attempt to parse a batch JSON array, with aggressive fallback. */
function tryParseBatchJson(text, expectedCount) {
  if (!text) return [];
  Logger.log("Parsing text of length: " + text.length);

  // 1. Try standard parse for the whole text (might be an array or an object with "results")
  try {
    const firstBracket = text.indexOf("[");
    const lastBracket = text.lastIndexOf("]");
    const firstBrace = text.indexOf("{");
    const lastBrace = text.lastIndexOf("}");

    // If it looks like an array, try parsing it directly
    if (firstBracket !== -1 && (firstBrace === -1 || firstBracket < firstBrace)) {
       const arrayText = text.substring(firstBracket, lastBracket + 1);
       const parsed = JSON.parse(arrayText);
       if (Array.isArray(parsed)) return parsed;
    }

    // If it looks like an object, try parsing it and looking for "results"
    if (firstBrace !== -1) {
      const objectText = text.substring(firstBrace, lastBrace + 1);
      const parsed = JSON.parse(objectText);
      if (parsed.results && Array.isArray(parsed.results)) return parsed.results;
      // If it's just a single object (happens sometimes with small chunks), wrap it
      if (parsed.question_id) return [parsed];
    }
  } catch (e) {
    Logger.log("Standard parse failed, trying aggressive extraction...");
  }

  // 2. Aggressive extraction using regex to find objects
  const results = [];
  // Look for anything that starts with { and has "question_id" followed by value
  const regex = /\{\s*"question_id"\s*:\s*"([^"]+)"/g;
  let match;
  
  while ((match = regex.exec(text)) !== null) {
    let pos = match.index;
    let braceCount = 0;
    let endPos = -1;
    
    for (let i = pos; i < text.length; i++) {
      if (text[i] === '{') braceCount++;
      else if (text[i] === '}') braceCount--;
      
      if (braceCount === 0) {
        endPos = i;
        break;
      }
    }
    
    if (endPos !== -1) {
      const candidate = text.substring(pos, endPos + 1);
      try {
        results.push(JSON.parse(candidate));
      } catch (e) {
        // Even more aggressive regex fallback for this specific fragment
        const qId = match[1];
        const scoreMatch = candidate.match(/"total_score"\s*:\s*(\d+(\.\d+)?)/);
        const feedbackMatch = candidate.match(/"feedback"\s*:\s*"([^"]+)"/);
        const improvementMatch = candidate.match(/"improvement"\s*:\s*"([^"]+)"/);
        
        results.push({
          question_id: qId,
          total_score: scoreMatch ? parseFloat(scoreMatch[1]) : 0,
          feedback: feedbackMatch ? feedbackMatch[1] : "Extracted via regex fallback.",
          improvement: improvementMatch ? improvementMatch[1] : "N/A",
          criteria_results: []
        });
      }
    }
  }

  Logger.log(`Extraction complete. Found ${results.length} questions.`);
  return results;
}



/**
 * Call Groq API (Llama 3.3 70B)
 */
function callGroqApi(prompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GROQ_API_KEY');
  if (!apiKey) throw new Error('GROQ_API_KEY not set in Script Properties.');

  const modelId = "llama-3.3-70b-versatile";
  const url = "https://api.groq.com/openai/v1/chat/completions";

  const payload = {
    model: modelId,
    messages: [
      { role: "system", content: "You are a strict exam grader. Always return valid JSON only." },
      { role: "user", content: prompt }
    ],
    temperature: 0.1,
    max_tokens: 2048,
    response_format: { type: "json_object" }
  };

  const params = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const resp = UrlFetchApp.fetch(url, params);
  const code = resp.getResponseCode();
  const text = resp.getContentText();

  Logger.log("Groq API response code: " + code);

  if (code !== 200) {
    throw new Error(`Groq API error ${code}: ${text}`);
  }

  const json = JSON.parse(text);
  if (!json.choices || !json.choices[0].message || !json.choices[0].message.content) {
    throw new Error("Groq response shape unexpected: " + text);
  }

  return json.choices[0].message.content;
}

/** Build HTML email body with professional styling. */
function buildEmailHtml(examName, studentName, results) {
  const totalScore = results.reduce((sum, r) => sum + (r.score || 0), 0);
  
  let rowsHtml = results.map(r => `
    <div style="margin-bottom: 25px; padding: 15px; border-left: 4px solid #4a90e2; background-color: #f9f9f9; border-radius: 4px;">
      <h3 style="margin-top: 0; color: #333; font-size: 16px;">Question: ${r.qHeader}</h3>
      <div style="display: flex; align-items: center; margin-bottom: 10px;">
        <span style="background-color: #4a90e2; color: white; padding: 4px 10px; border-radius: 20px; font-weight: bold; font-size: 14px;">Score: ${r.score}</span>
      </div>
      <p style="margin: 5px 0;"><strong>Feedback:</strong> ${r.feedback}</p>
      
      <div style="margin: 10px 0; padding-left: 15px; border-left: 2px solid #ddd;">
        <p style="margin: 0 0 5px 0; font-size: 13px; color: #666; font-weight: bold;">Criteria Breakdown:</p>
        <ul style="margin: 0; padding-left: 20px; font-size: 13px; color: #555;">
          ${(r.criteriaResults || []).map(c => `
            <li style="margin-bottom: 3px;">
              <strong>${c.id}:</strong> ${c.awarded} marks - <span style="font-style: italic;">${c.reason}</span>
            </li>
          `).join("")}
        </ul>
      </div>
      
      <p style="margin: 10px 0 0 0; color: #2c3e50; font-size: 14px;"><strong>💡 Improvement Idea:</strong> ${r.improvement}</p>
    </div>
  `).join("");

  return `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; color: #444; line-height: 1.6;">
      <div style="background-color: #2c3e50; color: white; padding: 30px 20px; text-align: center; border-radius: 8px 8px 0 0;">
        <h1 style="margin: 0; font-size: 24px;">Exam Results</h1>
        <p style="margin: 10px 0 0 0; opacity: 0.9;">${examName}</p>
      </div>
      
      <div style="padding: 20px; border: 1px solid #eee; border-top: none; border-radius: 0 0 8px 8px;">
        <p style="font-size: 18px; margin-top: 0;">Dear <strong>${studentName}</strong>,</p>
        <p>Congratulations on completing the examination! Here is the detailed breakdown of your performance as evaluated by our AI grading system.</p>
        
        <div style="background-color: #ebf5fb; padding: 15px; border-radius: 6px; text-align: center; margin: 20px 0;">
          <span style="font-size: 14px; color: #5dade2; text-transform: uppercase; font-weight: bold; letter-spacing: 1px;">Overall Performance</span>
          <div style="font-size: 32px; font-weight: bold; color: #2e86c1; margin-top: 5px;">Total Score: ${totalScore}</div>
        </div>

        ${rowsHtml}

        <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #eee; text-align: center; font-size: 12px; color: #999;">
          <p>This is an automated grade report based on predefined criteria. If you have any questions, please contact Yared - 0922761594.</p>
          <p>&copy; ${new Date().getFullYear()} Yared Technology School</p>
        </div>
      </div>
    </div>
  `;
}
