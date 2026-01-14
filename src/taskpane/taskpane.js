
// Constants for API endpoints (can be made configurable later if needed)
const OPENAI_API_ENDPOINT = "https://api.openai.com/v1/chat/completions";

const translations = {
  "Chinese": {
    "title": "AI 翻譯助手",
    "originalContentLabel": "信件內容同步 (原文)",
    "translateButton": "使用 GPT-4o-mini 翻譯",
    "translationResultLabel1": "翻譯結果 1 (GPT-4o-mini)",
    "translationResultLabel2": "翻譯結果 2 (Claude)",
    "insertButton1": "插入翻譯結果 1 至郵件",
    "insertButton2": "插入翻譯結果 2 至郵件",
    "originalTextPlaceholder": "正在自動同步郵件內容...",
    "translatedTextPlaceholder": "翻譯結果會顯示在這裡...",
    "langJapanese": "日文",
    "langEnglish": "英文",
    "langChinese": "中文",
    "statusTranslating": "正在翻譯中...",
    "statusSuccess": "翻譯完成。",
    "statusFail": "翻譯失敗，請檢查 API Key。",
    "statusInsertSuccess": "已成功插入郵件。",
    "statusInsertFail": "插入失敗。",
    "statusEnterContent": "請先輸入郵件內容。",
    "statusMissingApiKey": "請填寫所有必要的 API Key。",
    "statusMissingAtLeastOneApiKey": "請至少填寫一個 API Key。"
  },
  "Japanese": {
    "title": "AI 翻訳アシスタント",
    "originalContentLabel": "メール内容同期 (原文)",
    "translateButton": "GPT-4o-miniで翻訳",
    "translationResultLabel1": "翻訳結果 1 (GPT-4o-mini)",
    "translationResultLabel2": "翻訳結果 2 (Claude)",
    "insertButton1": "翻訳結果 1 をメールに挿入",
    "insertButton2": "翻訳結果 2 をメールに挿入",
    "originalTextPlaceholder": "メール内容を自動同期中...",
    "translatedTextPlaceholder": "翻訳結果はここに表示されます...",
    "langJapanese": "日本語",
    "langEnglish": "英語",
    "langChinese": "中国語",
    "statusTranslating": "翻訳中...",
    "statusSuccess": "翻訳完了。",
    "statusFail": "翻訳に失敗しました。APIキーを確認してください。",
    "statusInsertSuccess": "メールへの挿入に成功しました。",
    "statusInsertFail": "挿入に失敗しました。",
    "statusEnterContent": "まずメール内容を入力してください。",
    "statusMissingApiKey": "すべての必要なAPIキーを入力してください。",
    "statusMissingAtLeastOneApiKey": "少なくとも1つのAPIキーを入力してください。"
  },
  "English": {
    "title": "AI Translation Assistant",
    "originalContentLabel": "Email Content Sync (Original)",
    "translateButton": "Translate with GPT-4o-mini",
    "translationResultLabel1": "Translation Result 1 (GPT-4o-mini)",
    "translationResultLabel2": "Translation Result 2 (Claude)",
    "insertButton1": "Insert Translation 1 into Email",
    "insertButton2": "Insert Translation 2 into Email",
    "originalTextPlaceholder": "Syncing email content automatically...",
    "translatedTextPlaceholder": "Translation result will be displayed here...",
    "langJapanese": "Japanese",
    "langEnglish": "English",
    "langChinese": "Chinese",
    "statusTranslating": "Translating...",
    "statusSuccess": "Translation complete.",
    "statusFail": "Translation failed. Please check your API key.",
    "statusInsertSuccess": "Successfully inserted into email.",
    "statusInsertFail": "Insertion failed.",
    "statusEnterContent": "Please enter email content first.",
    "statusMissingApiKey": "Please fill in all required API Keys.",
    "statusMissingAtLeastOneApiKey": "Please fill in at least one API Key."
  }
};

function updateUIText(language) {
  const lang = translations[language] || translations["English"];

  document.title = lang.title;
  document.getElementById("originalContentLabel").innerText = lang.originalContentLabel;
  document.getElementById("translateBtn").innerText = lang.translateButton;
  document.getElementById("translationResultLabel1").innerText = lang.translationResultLabel1;
  document.getElementById("translationResultLabel2").innerText = lang.translationResultLabel2;
  document.getElementById("insertBtn").innerText = lang.insertButton1;
  document.getElementById("insertBtn2").innerText = lang.insertButton2;
  document.getElementById("originalText").placeholder = lang.originalTextPlaceholder;
  document.getElementById("translatedText").placeholder = lang.translatedTextPlaceholder;
  document.getElementById("translatedText2").placeholder = lang.translatedTextPlaceholder;

  const langSelector = document.getElementById("languageSelector");
  langSelector.querySelector('option[value="Japanese"]').innerText = lang.langJapanese;
  langSelector.querySelector('option[value="English"]').innerText = lang.langEnglish;
  langSelector.querySelector('option[value="Chinese"]').innerText = lang.langChinese;
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("translateBtn").onclick = translate;
    document.getElementById("insertBtn").onclick = insertTranslation1;
    document.getElementById("insertBtn2").onclick = insertTranslation2;
    document.getElementById("languageSelector").onchange = () => {
      const selectedLanguage = document.getElementById("languageSelector").value;
      updateUIText(selectedLanguage);
    };

    // Set initial UI text
    const initialLanguage = document.getElementById("languageSelector").value;
    updateUIText(initialLanguage);

    // Load API keys from roaming settings
    loadApiKeys();

    // 每 2 秒自動同步一次內容
    setInterval(syncMailBody, 2000);
  }
});

/** Load API keys from roaming settings */
function loadApiKeys() {
  const openaiApiKey = Office.context.roamingSettings.get('openaiApiKey');
  if (openaiApiKey) {
    document.getElementById('openaiApiKey').value = openaiApiKey;
  }
  const secondLLMApiKey = Office.context.roamingSettings.get('secondLLMApiKey');
  if (secondLLMApiKey) {
    document.getElementById('secondLLMApiKey').value = secondLLMApiKey;
  }
}

/** Save API keys to roaming settings */
function saveApiKeys(openaiApiKey, secondLLMApiKey) {
  Office.context.roamingSettings.set('openaiApiKey', openaiApiKey);
  Office.context.roamingSettings.set('secondLLMApiKey', secondLLMApiKey);
  Office.context.roamingSettings.saveAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.error('Failed to save API keys:', result.error.message);
    } else {
      // Keys saved successfully, no need to bother the user with a message.
      console.log('API keys saved.');
    }
  });
}

/** 1. 同步郵件內容 */
async function syncMailBody() {
  // 檢查是否在編輯模式 (Compose mode)
  if (Office.context.mailbox.item.body) {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const originalBox = document.getElementById("originalText");
        // 只有在內容有變動時才更新，避免干擾使用者操作
        if (originalBox.value !== result.value) {
          originalBox.value = result.value;
        }
      }
    });
  }
}

/** 2. 呼叫 GPT-4o-mini & Claude 翻譯 */
async function translate() {
  const text = document.getElementById("originalText").value;
  const status = document.getElementById("status");
  const selectedLanguage = document.getElementById("languageSelector").value;
  const lang = translations[selectedLanguage] || translations["English"];

  const openaiApiKey = document.getElementById("openaiApiKey").value;
  const secondLLMApiKey = document.getElementById("secondLLMApiKey").value;

  // Clear previous results
  document.getElementById("translatedText").value = "";
  document.getElementById("translatedText2").value = "";
  document.getElementById("insertBtn").style.display = "none";
  document.getElementById("insertBtn2").style.display = "none";

  if (!text) {
    status.innerText = lang.statusEnterContent;
    return;
  }

  // Save any keys that are present
  if (openaiApiKey || secondLLMApiKey) {
    saveApiKeys(openaiApiKey, secondLLMApiKey);
  } else {
    status.innerText = lang.statusMissingAtLeastOneApiKey;
    return;
  }

  status.innerText = lang.statusTranslating;

  let prompt1 = ""; // For GPT-4o-mini
  let prompt2 = ""; // For Claude (simulated)

  switch (selectedLanguage) {
    case "Japanese":
      prompt1 = "你是一個專業的醫藥與商業翻譯專家。請將以下中文內容翻譯成流暢的日文。直接回傳翻譯結果即可。";
      prompt2 = "你是頂尖的創意翻譯家，請將下列內容翻譯成自然且帶有豐富情感的日文。";
      break;
    case "English":
      prompt1 = "You are a professional medical and business translator. Please translate the following Chinese content into fluent English. Return the translation directly.";
      prompt2 = "You are a top creative translator. Translate the following content into natural and emotionally rich English.";
      break;
    case "Chinese":
      prompt1 = "你是一個專業的醫藥與商業翻譯專家。請將以下內容翻譯成通順的中文。直接回傳翻譯結果即可。";
      prompt2 = "你是頂尖的創意翻譯家，請將下列內容翻譯成自然且帶有豐富情感的中文。";
      break;
  }

  const promises = [];
  const uiTargets = [];

  if (openaiApiKey) {
    promises.push(callOpenAI(openaiApiKey, prompt1, text, "gpt-4o-mini").catch(e => ({ error: true, message: `翻譯 1 失敗: ${e.message}` })));
    uiTargets.push({ textAreaId: 'translatedText', buttonId: 'insertBtn' });
  }

  if (secondLLMApiKey) {
    promises.push(callOpenAI(secondLLMApiKey, prompt2, text, "gpt-4o-mini").catch(e => ({ error: true, message: `翻譯 2 失敗: ${e.message}` })));
    uiTargets.push({ textAreaId: 'translatedText2', buttonId: 'insertBtn2' });
  }

  try {
    const results = await Promise.all(promises);
    let allSucceeded = true;
    let errorMessages = [];

    results.forEach((result, index) => {
      const target = uiTargets[index];
      if (result.error) {
        allSucceeded = false;
        errorMessages.push(result.message);
        document.getElementById(target.textAreaId).value = result.message; // Display error in textarea
      } else {
        document.getElementById(target.textAreaId).value = result;
        // document.getElementById(target.buttonId).style.display = 'block'; // Keep insert buttons hidden as per user request
      }
    });

    if (allSucceeded) {
      status.innerText = lang.statusSuccess;
    } else {
      status.innerText = errorMessages.join('; ');
    }

  } catch (error) { // Should not happen with individual catch blocks, but as a fallback
    console.error(error);
    status.innerText = "發生意外錯誤: " + error.message;
  }
}

async function callOpenAI(apiKey, prompt, text, model) {
  const response = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model: model,
      messages: [
        { role: "system", content: prompt },
        { role: "user", content: text }
      ],
      temperature: 0.3
    })
  });

  const data = await response.json();
  if (response.ok && data.choices && data.choices.length > 0) {
    return data.choices[0].message.content;
  } else {
    const errorMessage = data.error?.message || "API request failed.";
    throw new Error(errorMessage);
  }
}


/** 3. 插入文字到郵件游標處 */
function insertTranslation1() {
  const translatedText = document.getElementById("translatedText").value;
  insertTextIntoMail(translatedText);
}

function insertTranslation2() {
  const translatedText = document.getElementById("translatedText2").value;
  insertTextIntoMail(translatedText);
}

function insertTextIntoMail(textToInsert) {
    const selectedLanguage = document.getElementById("languageSelector").value;
    const lang = translations[selectedLanguage] || translations["English"];
    
    if (!textToInsert) return;
  
    Office.context.mailbox.item.body.setSelectedDataAsync(
        textToInsert,
      { coercionType: Office.CoercionType.Text },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          document.getElementById("status").innerText = lang.statusInsertSuccess;
        } else {
          document.getElementById("status").innerText = lang.statusInsertFail;
        }
      }
    );
}
