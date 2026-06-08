const OWNER_EMAIL = "revax.int.works@gmail.com";
const FORM_TITLE = "インテリアREVAX お問い合わせフォーム";
const SHEET_TITLE = "インテリアREVAX お問い合わせ管理";

function setupRevaxGoogleInquiryTool() {
  const spreadsheet = SpreadsheetApp.create(SHEET_TITLE);
  const form = FormApp.create(FORM_TITLE);

  form.setDescription("クロス、CF、フロアタイル、内装リフォームのご相談はこちらからお送りください。");
  form.setCollectEmail(false);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());

  form.addTextItem().setTitle("お名前 / Name / 姓名").setRequired(true);
  form.addTextItem().setTitle("メールアドレス / Email / 邮箱").setRequired(true);
  form.addTextItem().setTitle("電話番号 / Phone / 电话").setRequired(false);
  form.addMultipleChoiceItem()
    .setTitle("希望言語 / Preferred language / 希望语言")
    .setChoiceValues(["日本語", "English", "中文"])
    .setRequired(true);
  form.addCheckboxItem()
    .setTitle("ご相談内容 / Inquiry type / 咨询内容")
    .setChoiceValues(["クロス工事", "CF施工", "フロアタイル施工", "内装リフォーム", "原状回復", "その他"])
    .setRequired(true);
  form.addTextItem().setTitle("施工場所・エリア / Location / 施工地点").setRequired(false);
  form.addTextItem().setTitle("希望時期 / Preferred timing / 希望时间").setRequired(false);
  form.addParagraphTextItem()
    .setTitle("詳細 / Details / 详细内容")
    .setHelpText("現場の状況、広さ、気になる箇所などをご記入ください。")
    .setRequired(true);

  const summarySheet = spreadsheet.insertSheet("問い合わせ整理");
  summarySheet.appendRow([
    "受付日時",
    "お名前",
    "メール",
    "電話",
    "希望言語",
    "相談内容",
    "施工場所",
    "希望時期",
    "原文",
    "日本語訳",
    "確認ポイント",
    "返信案",
    "Gmail下書き",
  ]);
  summarySheet.setFrozenRows(1);
  summarySheet.autoResizeColumns(1, 13);

  const properties = PropertiesService.getScriptProperties();
  properties.setProperty("SUMMARY_SHEET_ID", spreadsheet.getId());
  properties.setProperty("FORM_ID", form.getId());

  ScriptApp.newTrigger("handleRevaxFormSubmit")
    .forForm(form)
    .onFormSubmit()
    .create();

  Logger.log("フォームURL: " + form.getPublishedUrl());
  Logger.log("編集用フォームURL: " + form.getEditUrl());
  Logger.log("管理スプレッドシートURL: " + spreadsheet.getUrl());
}

function handleRevaxFormSubmit(e) {
  const values = getResponseValues_(e.response);
  const submittedAt = new Date();
  const name = values["お名前 / Name / 姓名"] || "";
  const email = values["メールアドレス / Email / 邮箱"] || "";
  const phone = values["電話番号 / Phone / 电话"] || "";
  const language = values["希望言語 / Preferred language / 希望语言"] || "";
  const inquiryType = values["ご相談内容 / Inquiry type / 咨询内容"] || "";
  const location = values["施工場所・エリア / Location / 施工地点"] || "";
  const timing = values["希望時期 / Preferred timing / 希望时间"] || "";
  const details = values["詳細 / Details / 详细内容"] || "";

  const translatedDetails = translateToJapanese_(details, language);
  const checkpoints = buildCheckpoints_(location, timing, details);
  const replyDraft = buildReplyDraft_(name, language, inquiryType, location, timing);
  const draftStatus = createGmailDraft_(email, replyDraft.subject, replyDraft.body);

  const spreadsheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty("SUMMARY_SHEET_ID"));
  const sheet = spreadsheet.getSheetByName("問い合わせ整理");
  sheet.appendRow([
    submittedAt,
    name,
    email,
    phone,
    language,
    inquiryType,
    location,
    timing,
    details,
    translatedDetails,
    checkpoints,
    replyDraft.body,
    draftStatus,
  ]);
  sheet.autoResizeColumns(1, 13);

  const ownerSubject = "【インテリアREVAX】新しいお問い合わせ";
  const ownerBody = [
    "新しいお問い合わせが届きました。",
    "",
    "【お名前】" + name,
    "【メール】" + email,
    "【電話】" + phone,
    "【希望言語】" + language,
    "【相談内容】" + inquiryType,
    "【施工場所】" + location,
    "【希望時期】" + timing,
    "",
    "【原文】",
    details,
    "",
    "【日本語訳】",
    translatedDetails,
    "",
    "【確認ポイント】",
    checkpoints,
    "",
    "【返信案】",
    replyDraft.body,
    "",
    "管理シート: " + spreadsheet.getUrl(),
  ].join("\n");

  GmailApp.sendEmail(OWNER_EMAIL, ownerSubject, ownerBody);
}

function getResponseValues_(response) {
  const values = {};
  response.getItemResponses().forEach((itemResponse) => {
    const title = itemResponse.getItem().getTitle();
    const answer = itemResponse.getResponse();
    values[title] = Array.isArray(answer) ? answer.join(", ") : answer;
  });
  return values;
}

function translateToJapanese_(text, language) {
  if (!text) return "";
  if (language === "日本語") return text;

  const source = language === "English" ? "en" : language === "中文" ? "zh" : "auto";
  try {
    return LanguageApp.translate(text, source, "ja");
  } catch (error) {
    return "翻訳できませんでした。原文を確認してください。\n\n" + text;
  }
}

function buildCheckpoints_(location, timing, details) {
  const points = [];
  if (!location) points.push("施工場所・住所の確認");
  if (!timing) points.push("現地確認または施工希望時期の確認");
  points.push("施工範囲、部屋数、広さの確認");
  points.push("現在の状態が分かる写真の有無を確認");
  points.push("家具移動や駐車スペースの有無を確認");
  if (details && details.length < 30) points.push("詳細内容が少ないため、追加ヒアリングが必要");
  return points.map((point) => "・" + point).join("\n");
}

function buildReplyDraft_(name, language, inquiryType, location, timing) {
  const displayName = name || "お問い合わせのお客様";
  const locationLine = location ? "施工場所は「" + location + "」で承りました。" : "施工場所についても確認させてください。";
  const timingLine = timing ? "希望時期は「" + timing + "」で承りました。" : "ご希望の時期があればお知らせください。";

  if (language === "English") {
    return {
      subject: "Thank you for contacting Interior REVAX",
      body: [
        "Dear " + displayName + ",",
        "",
        "Thank you for contacting Interior REVAX.",
        "We received your inquiry about " + inquiryType + ".",
        "",
        "To prepare an estimate, please send us the site address, approximate room size or work area, preferred date for a site check, and photos if available.",
        "",
        "We will review the details and reply with the next steps.",
        "",
        "Interior REVAX",
      ].join("\n"),
    };
  }

  if (language === "中文") {
    return {
      subject: "感谢您联系 Interior REVAX",
      body: [
        displayName + " 您好：",
        "",
        "感谢您联系 Interior REVAX。",
        "我们已收到关于「" + inquiryType + "」的咨询。",
        "",
        "为了确认报价，请告知施工地址、房间大小或施工范围、希望现场确认的日期，如有照片也请一并发送。",
        "",
        "确认内容后，我们会再联系您说明下一步。",
        "",
        "Interior REVAX",
      ].join("\n"),
    };
  }

  return {
    subject: "お問い合わせありがとうございます｜インテリアREVAX",
    body: [
      displayName + " 様",
      "",
      "お問い合わせありがとうございます。",
      "インテリアREVAXです。",
      "",
      "「" + inquiryType + "」についてご相談を承りました。",
      locationLine,
      timingLine,
      "",
      "お見積りのため、施工住所、施工範囲やお部屋の広さ、現地確認のご希望日をお知らせください。",
      "可能でしたら、現在の状態が分かるお写真もお送りいただけますとスムーズです。",
      "",
      "内容を確認後、必要な工事と費用の目安をご案内いたします。",
      "",
      "インテリアREVAX",
    ].join("\n"),
  };
}

function createGmailDraft_(to, subject, body) {
  if (!to) return "メールアドレス未入力のため下書き未作成";
  try {
    GmailApp.createDraft(to, subject, body);
    return "Gmail下書き作成済み";
  } catch (error) {
    return "Gmail下書き作成失敗: " + error.message;
  }
}
