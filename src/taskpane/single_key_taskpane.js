
// Constants for API endpoints (can be made configurable later if needed)
const OPENAI_API_ENDPOINT = "https://api.openai.com/v1/chat/completions";

const translations = {
  "Chinese": {
    "title": "AI 翻譯助手",
    "uiLanguageSection": "UI 語言",
    "apiKeySectionTitle": "請填寫 API Key (避免硬寫在程式碼中)",
    "openaiApiKeyLabel": "OpenAI API Key (用於 GPT-4o-mini):",
    "originalContentLabel": "信件內容同步 (原文)",
    "translationLanguageLabel": "選擇翻譯語言",
    "translateButton": "使用 GPT-4o-mini 翻譯",
    "translationResultLabel1": "翻譯結果 (GPT-4o-mini)",
    "insertButton1": "插入翻譯結果至郵件",
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
    "statusMissingApiKey": "請填寫 API Key。"
  },
  "Japanese": {
    "title": "AI 翻訳アシスタント",
    "uiLanguageSection": "UI言語",
    "apiKeySectionTitle": "APIキーを入力してください（コードにハードコーディングしないでください）",
    "openaiApiKeyLabel": "OpenAI APIキー (GPT-4o-mini用):",
    "originalContentLabel": "メール内容同期 (原文)",
    "translationLanguageLabel": "翻訳言語を選択",
    "translateButton": "GPT-4o-miniで翻訳",
    "translationResultLabel1": "翻訳結果 (GPT-4o-mini)",
    "insertButton1": "翻訳結果をメールに挿入",
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
    "statusMissingApiKey": "APIキーを入力してください。"
  },
  "English": {
    "title": "AI Translation Assistant",
    "uiLanguageSection": "UI Language",
    "apiKeySectionTitle": "Please fill in API Key (avoid hardcoding in the source code)",
    "openaiApiKeyLabel": "OpenAI API Key (for GPT-4o-mini):",
    "originalContentLabel": "Email Content Sync (Original)",
    "translationLanguageLabel": "Select Translation Language",
    "translateButton": "Translate with GPT-4o-mini",
    "translationResultLabel1": "Translation Result (GPT-4o-mini)",
    "insertButton1": "Insert Translation into Email",
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
    "statusMissingApiKey": "Please fill in the API Key."
  }
};

function updateUIText(language) {
  const lang = translations[language] || translations["English"];

  document.title = lang.title;
  document.getElementById("uiLanguageSection").innerText = lang.uiLanguageSection;
  document.getElementById("apiKeySectionTitle").innerText = lang.apiKeySectionTitle;
  document.getElementById("openaiApiKeyLabel").innerText = lang.openaiApiKeyLabel;
  document.getElementById("originalContentLabel").innerText = lang.originalContentLabel;
  document.getElementById("translationLanguageLabel").innerText = lang.translationLanguageLabel;
  document.getElementById("translateBtn").innerText = lang.translateButton;
  document.getElementById("translationResultLabel1").innerText = lang.translationResultLabel1;
  document.getElementById("insertBtn").innerText = lang.insertButton1;
  document.getElementById("originalText").placeholder = lang.originalTextPlaceholder;
  document.getElementById("translatedText").placeholder = lang.translatedTextPlaceholder;

  const langSelector = document.getElementById("languageSelector");
  langSelector.querySelector('option[value="Japanese"]').innerText = lang.langJapanese;
  langSelector.querySelector('option[value="English"]').innerText = lang.langEnglish;
  langSelector.querySelector('option[value="Chinese"]').innerText = lang.langChinese;
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("translateBtn").onclick = translate;
    document.getElementById("insertBtn").onclick = insertTranslation;
    document.getElementById("uiLanguageSelector").onchange = () => {
      const selectedLanguage = document.getElementById("uiLanguageSelector").value;
      updateUIText(selectedLanguage);
    };

    // Set initial UI text
    const initialLanguage = document.getElementById("uiLanguageSelector").value;
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
}

/** Save API keys to roaming settings */
function saveApiKeys(openaiApiKey) {
  Office.context.roamingSettings.set('openaiApiKey', openaiApiKey);
  Office.context.roamingSettings.saveAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.error('Failed to save API keys:', result.error.message);
    } else {
      console.log('API keys saved.');
    }
  });
}

/** 1. 同步郵件內容 */
async function syncMailBody() {
  if (Office.context.mailbox.item.body) {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const originalBox = document.getElementById("originalText");
        if (originalBox.value !== result.value) {
          originalBox.value = result.value;
        }
      }
    });
  }
}

/** 2. 呼叫 GPT-4o-mini 翻譯 */
async function translate() {
  const text = document.getElementById("originalText").value;
  const status = document.getElementById("status");
  const selectedUiLanguage = document.getElementById("uiLanguageSelector").value;
  const lang = translations[selectedUiLanguage] || translations["English"];

  const openaiApiKey = document.getElementById("openaiApiKey").value;

  document.getElementById("translatedText").value = "";
  document.getElementById("insertBtn").style.display = "none";

  if (!text) {
    status.innerText = lang.statusEnterContent;
    return;
  }

  if (openaiApiKey) {
    saveApiKeys(openaiApiKey);
  } else {
    status.innerText = lang.statusMissingApiKey;
    return;
  }

  status.innerText = lang.statusTranslating;

  let prompt = "";
  const selectedLanguage = document.getElementById("languageSelector").value;

  switch (selectedLanguage) {
    case "Japanese":
      prompt = "你是一個專業的醫藥與商業翻譯專家。請將以下中文內容翻譯成流暢的日文。直接回傳翻譯結果即可。";
      break;
    case "English":
      prompt = "You are a professional medical and business translator. Please translate the following Chinese content into fluent English. Return the translation directly.";
      break;
    case "Chinese":
      prompt = "你是一個專業的醫藥與商業翻譯專家。請將以下內容翻譯成通順的中文。直接回傳翻譯結果即可。";
      break;
  }

  try {
    const result = await callOpenAI(openaiApiKey, prompt, text, "gpt-4o-mini");
    document.getElementById("translatedText").value = result;
    document.getElementById("insertBtn").style.display = 'block';
    status.innerText = lang.statusSuccess;
  } catch (error) {
    console.error(error);

    document.getElementById("translatedText").value = error.message;
    status.innerText = lang.statusFail;
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
function insertTranslation() {
  const translatedText = document.getElementById("translatedText").value;
  insertTextIntoMail(translatedText);
}

function insertTextIntoMail(textToInsert) {
    const selectedUiLanguage = document.getElementById("uiLanguageSelector").value;
    const lang = translations[selectedUiLanguage] || translations["English"];
    
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
