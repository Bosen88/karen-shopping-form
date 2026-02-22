import { GoogleSpreadsheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';

// =========================================
// 1. 解析：把 "所有勾選項目" 解析為結構化欄位
// =========================================
function parseInterestsToStructured(allInterests) {
  var raw = allInterests || [];
  var items = [];
  if (Array.isArray(raw)) {
    items = raw.slice();
  } else {
    items = String(raw).split(/[,;\/\|，；]/).map(function(s){ return s.trim(); }).filter(Boolean);
  }

  var synonymMap = {
    skin: ["中性肌","乾性肌","油性肌","混合性","敏感肌","dry","oily","normal","combination","sensitive"],
    skincare: ["保濕","美白","淡斑","粉刺","痘痘","抗老","除皺","防曬","隔離","潔膚乳","化妝水","乳液","面膜","moisturize","whitening","anti-age"],
    hair: ["頭皮","頭皮屑","頭皮癢","掉髮","髮量稀疏","染燙","乾燥","毛躁","出油","頭皮去角質","洗髮","潤髮","造型","hair","dandruff","hair loss"],
    cleaning: ["萬用清潔","碗盤","洗潔","去油","地毯","不鏽鋼","花崗石","磁磚","清潔用品組","衛生紙","廚房紙巾","clean","detergent"],
    nutrition: ["益生菌","B群","葉黃素","膠原蛋白","魚油","q10","體重管理","代餐","營養","消化","腸胃","睡眠","壓力","probiotic","collagen","omega","vitamin"],
    travel: ["旅行","國內","國外","klook","kkday","agoda","booking"]
  };

  var result = { skinTypes: [], skincareNeeds: [], hairNeeds: [], cleaningNeeds: [], nutritionNeeds: [], travelNeeds: [], otherNeeds: [] };

  function normalize(s) { return String(s || '').trim().toLowerCase(); }

  var index = {};
  for (var cat in synonymMap) {
    synonymMap[cat].forEach(function(k) { index[normalize(k)] = cat; });
  }

  items.forEach(function(it) {
    var ni = normalize(it);
    var placed = false;
    for (var kw in index) {
      if (ni.indexOf(kw) !== -1) {
        var cat = index[kw];
        if (cat === 'skin') { if (result.skinTypes.indexOf(it) === -1) result.skinTypes.push(it); }
        else if (cat === 'skincare') { if (result.skincareNeeds.indexOf(it) === -1) result.skincareNeeds.push(it); }
        else if (cat === 'hair') { if (result.hairNeeds.indexOf(it) === -1) result.hairNeeds.push(it); }
        else if (cat === 'cleaning') { if (result.cleaningNeeds.indexOf(it) === -1) result.cleaningNeeds.push(it); }
        else if (cat === 'nutrition') { if (result.nutritionNeeds.indexOf(it) === -1) result.nutritionNeeds.push(it); }
        else if (cat === 'travel') { if (result.travelNeeds.indexOf(it) === -1) result.travelNeeds.push(it); }
        placed = true;
        break;
      }
    }
    if (!placed) { if (result.otherNeeds.indexOf(it) === -1) result.otherNeeds.push(it); }
  });

  Object.keys(result).forEach(function(k) { result[k].sort(); });
  return result;
}

// =========================================
// 2. 計分邏輯
// =========================================
function computeCustomerScore(parsed, formData) {
  var score = 0;
  var rep = (formData.repurchaseFrequency || '').toString();
  if (rep.indexOf('每周') !== -1 || rep.indexOf('每週') !== -1) score += 4;
  if (rep.indexOf('每月') !== -1) score += 3;
  if (rep.indexOf('每季') !== -1) score += 1;

  if ((formData.consentDiscount || '').indexOf('願意') !== -1) score += 1;
  if (parsed && parsed.nutritionNeeds && parsed.nutritionNeeds.length >= 1) score += 2;
  if (parsed && parsed.skincareNeeds && parsed.skincareNeeds.length >= 2) score += 2;

  var mainCatsCount = (formData.mainCategories && formData.mainCategories.length) ? formData.mainCategories.length : 0;
  score += Math.min(mainCatsCount, 3);
  return score;
}

// =========================================
// 3. 建立 LINE Flex Message (原汁原味)
// =========================================
function buildFlexOrTextMessages(formData, SHEET_VIEW_URL, SHOP_URL) {
  function displayText(s, n) {
    if (s === undefined || s === null) return "—";
    var t = String(s).trim();
    if (t === "") return "—";
    if (n && t.length > n) return t.substring(0, n-1) + "…";
    return t;
  }
  function safeTrunc(s, n) { if (!s && s !== 0) return ""; s = String(s); return s.length > n ? s.substring(0, n-1) + "…" : s; }
  function websitesForDisplay(w) { if (!w) return ""; w = String(w); return w.length > 36 ? w.substring(0,33) + "…" : w; }
  function toBulleted(arr, maxItems, maxChars) {
    if (!arr || arr.length === 0) return "無";
    var items = arr.slice(0, maxItems).map(function(x) {
      var s = String(x).trim();
      if (maxChars && s.length > maxChars) s = s.substring(0, maxChars - 1) + "…";
      return "• " + s;
    });
    if (arr.length > maxItems) items.push("• …等 " + arr.length + " 項");
    return items.join("\n");
  }

  try {
    var name = (formData && formData.name) ? String(formData.name) : "";
    var email = (formData && formData.email) ? String(formData.email) : "";
    var lineName = (formData && (formData.lineName_placeholder || formData.lineName)) ? String(formData.lineName_placeholder || formData.lineName) : "";
    var top3 = (formData && formData.top3Needs) ? String(formData.top3Needs) : "";
    var mainCats = (formData && formData.mainCategories && formData.mainCategories.length) ? formData.mainCategories : [];
    var interests = (formData && formData.allInterests && formData.allInterests.length) ? formData.allInterests : [];
    var habit = (formData && formData.habit) ? String(formData.habit) : "";
    var websites = (formData && formData.websites) ? String(formData.websites) : "";
    var reasons = (formData && formData.reasons && formData.reasons.length) ? formData.reasons : [];
    var opportunity = (formData && formData.opportunity) ? String(formData.opportunity) : "";
    var brandPref = (formData && formData.brandPref) ? String(formData.brandPref) : "";
    var repurchaseFrequency = (formData && formData.repurchaseFrequency) ? String(formData.repurchaseFrequency) : "";
    var consentDiscount = (formData && formData.consentDiscount) ? String(formData.consentDiscount) : "";
    var parsed = (formData && formData.parsed) ? formData.parsed : {};
    var score = (typeof formData.score !== 'undefined') ? formData.score : null;

    var mainCatsText = toBulleted(mainCats, 6, 24);
    var interestsText = toBulleted(interests, 8, 36);
    var reasonsText = toBulleted(reasons, 6, 36);

    // Node.js 產生台灣時間字串
    var fillTime = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' }).replace(/\//g, '-');

    var bubble = {
      type: "bubble",
      header: {
        type: "box", layout: "vertical", backgroundColor: "#0d6efd", paddingAll: "12px",
        contents: [{ type: "text", text: "顧客健診回覆（管理通知）", weight: "bold", size: "lg", color: "#ffffff" }]
      },
      body: {
        type: "box", layout: "vertical", spacing: "sm",
        contents: [
          { type: "text", text: "基本資料", weight: "bold", size: "sm", margin: "md" },
          { type: "box", layout: "baseline", contents: [
            { type:"text", text:"姓名", color:"#666666", size:"xs", flex:1 },
            { type:"text", text: displayText(name, 40), size:"xs", flex:3, wrap:true }
          ] },
          { type: "box", layout: "baseline", contents: [
            { type:"text", text:"Email", color:"#666666", size:"xs", flex:1 },
            { type:"text", text: displayText(email, 40), size:"xs", flex:3, wrap:true }
          ] },
          { type: "box", layout: "baseline", contents: [
            { type:"text", text:"LINE 暱稱", color:"#666666", size:"xs", flex:1 },
            { type:"text", text: displayText(lineName, 30), size:"xs", flex:3, wrap:true }
          ] },
          { type: "separator", margin: "md" },
          { type: "text", text: "優先三項", weight: "bold", size: "sm", margin: "md" },
          { type: "text", text: displayText(top3, 140), size: "sm", wrap: true, margin: "xs" },
          { type: "text", text: "主要分類", weight: "bold", size: "sm", margin: "md" },
          { type: "text", text: displayText(mainCatsText, 200), size: "sm", wrap: true, margin: "xs" },
          { type: "text", text: "興趣項目（細項）", weight: "bold", size: "sm", margin: "md" },
          { type: "text", text: displayText(interestsText, 300), size: "sm", wrap: true, margin: "xs" },
          { type: "text", text: "購物習慣 / 常逛平台", weight: "bold", size: "sm", margin: "md" },
          { type: "text", text: displayText(habit + "  ｜  " + websitesForDisplay(websites), 80), size: "xs", color:"#666666", wrap:true, margin:"xs" },
          { type: "text", text: "選購原因", size: "sm", color: "#666666", margin: "md" },
          { type: "text", text: displayText(reasonsText, 200), size: "sm", wrap: true, margin: "xs" },
          { type: "text", text: "品牌偏好：" + displayText(brandPref, 60), size: "xs", margin: "md" },
          { type: "text", text: "回購頻率：" + displayText(repurchaseFrequency, 30) + "  ｜  折扣通知：" + displayText(consentDiscount, 20), size: "xs", margin: "xs" },
          { type: "text", text: "創業意願：" + displayText(opportunity, 40), size: "xs", margin: "md" },
          { type: "separator", margin: "md" },
          { type: "text", text: "解析標籤（系統自動）", weight: "bold", size: "sm", margin: "md" },
          { type: "text", text:
              "膚質: " + (parsed.skinTypes ? parsed.skinTypes.join(', ') : "—") +
              "\n護膚: " + (parsed.skincareNeeds ? parsed.skincareNeeds.join(', ') : "—") +
              "\n頭髮: " + (parsed.hairNeeds ? parsed.hairNeeds.join(', ') : "—") +
              "\n保健: " + (parsed.nutritionNeeds ? parsed.nutritionNeeds.join(', ') : "—")
            , size: "xs", color: "#666666", wrap: true, margin: "xs" },
          { type: "text", text: "填表時間：" + fillTime, size: "xxs", color: "#999999", margin: "md" },
          { type: "text", text: "CustomerScore：" + (score !== null ? String(score) : "—"), size: "xxs", color: "#999999", margin: "xs" }
        ]
      },
      footer: {
        type: "box", layout: "vertical", spacing: "sm",
        contents: [
          { type: "button", style: "link", height: "sm", action: { type: "uri", label: "查看完整回覆（表單）", uri: (SHEET_VIEW_URL && SHEET_VIEW_URL.length) ? SHEET_VIEW_URL : "https://docs.google.com/spreadsheets/d/1E6C9Wdc7Ij22m44k-AoRdkSwnH2auXIQ_KQ_lMwS98A/edit?usp=sharing" } },
          { type: "button", style: "link", height: "sm", action: { type: "postback", label: "貼上私訊範本", data: "action=copy_template&email=" + encodeURIComponent(email) + "&name=" + encodeURIComponent(name), displayText: displayText("私訊範本：您好，感謝您填問卷！關於您優先的「" + safeTrunc(top3,40) + "」，我這邊有優惠，是否要我協助下單或寄試用？", 150) } },
          { type: "button", style: "primary", color: "#0d6efd", height: "sm", action: { type: "uri", label: "開啟 SHOP.COM", uri: (SHOP_URL && SHOP_URL.length) ? SHOP_URL : "https://tw.shop.com/AUREVOIR2047" } }
        ],
        flex: 0
      }
    };

    var approxLength = JSON.stringify(bubble).length;
    if (approxLength < 8000) {
      return [{ type: "flex", altText: "新居家購物清單回覆 - " + displayText(name, 30), contents: bubble }];
    } else {
      var shortText = "🎉 新回覆：" + displayText(name, 30) + "\nEmail：" + displayText(email, 40) + "\n優先：" + displayText(top3,60) + "\n查看：" + ((SHEET_VIEW_URL && SHEET_VIEW_URL.length) ? SHEET_VIEW_URL : "");
      return [{ type: "text", text: shortText }];
    }
  } catch (err) {
    return [{ type: "text", text: "新回覆：" + (formData && formData.name ? formData.name : "使用者") }];
  }
}

// =========================================
// 4. Vercel 主要執行區域
// =========================================
export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ success: false, message: 'Method Not Allowed' });

  try {
    const formData = req.body;

    // 🛡️ 終極防呆機制：自動清理 Vercel 環境變數的雜訊
    const clientEmail = (process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || '').replace(/"/g, '').trim();
    let privateKey = process.env.GOOGLE_PRIVATE_KEY || '';
    privateKey = privateKey.replace(/"/g, '').replace(/\\n/g, '\n').trim();
    const sheetId = (process.env.GOOGLE_SHEET_ID || '').replace(/"/g, '').trim();

    if (!clientEmail || !privateKey || !sheetId) {
        throw new Error('環境變數遺失：請檢查 Vercel 後台是否已設定 GOOGLE_SERVICE_ACCOUNT_EMAIL, GOOGLE_PRIVATE_KEY, GOOGLE_SHEET_ID');
    }

    // Google Sheets 驗證
    const serviceAccountAuth = new JWT({ email: clientEmail, key: privateKey, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
    const doc = new GoogleSpreadsheet(sheetId, serviceAccountAuth);
    await doc.loadInfo(); 
    
    const sheet = doc.sheetsByTitle['顧客輪廓分析'];
    if (!sheet) throw new Error('找不到名為 "顧客輪廓分析" 的工作表');

    // 執行分析與計分
    const allInterestsStr = formData.allInterests ? formData.allInterests.join(', ') : "";
    const parsed = parseInterestsToStructured(allInterestsStr);
    const score = computeCustomerScore(parsed, formData);

    // 寫入 Google Sheet (完整欄位，與你原本的 REQUIRED_HEADERS 對齊)
    const newRow = {
      'Email': formData.email || "未提供Email",
      '姓名': formData.name || "未提供姓名",
      'LINE暱稱': formData.lineName_placeholder || "",
      '生日-年': formData.birthYear || "",
      '生日-月': formData.birthMonth || "",
      '主要興趣類別': formData.mainCategories ? formData.mainCategories.join(', ') : "",
      '所有勾選項目': allInterestsStr,
      '膚質類型': parsed.skinTypes.join(', '),
      '護膚需求': parsed.skincareNeeds.join(', '),
      '頭皮/頭髮需求': parsed.hairNeeds.join(', '),
      '清潔用品需求': parsed.cleaningNeeds.join(', '),
      '保健類需求': parsed.nutritionNeeds.join(', '),
      '品牌偏好': formData.brandPref || "",
      '回購頻率': formData.repurchaseFrequency || "",
      '是否願意折扣通知': formData.consentDiscount || "",
      '最在意的三項': formData.top3Needs || "",
      '購物習慣': formData.habit || "",
      '常逛網站': formData.websites || "",
      '購物原因': formData.reasons ? formData.reasons.join(', ') : "",
      '對機會感興趣': formData.opportunity || "",
      'CustomerScore': score,
      '最後更新時間': new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' })
    };
    await sheet.addRow(newRow);

    // 把資料裝回 formData 讓 LINE 通知可以使用
    formData.parsed = parsed;
    formData.score = score;

    // 發送 LINE 通知 (呼叫我們剛剛移植過來的完整 Flex Message 產生器)
    const lineToken = (process.env.LINE_CHANNEL_ACCESS_TOKEN || '').replace(/"/g, '').trim();
    const lineIdsString = (process.env.LINE_RECIPIENT_USER_IDS || '').replace(/"/g, '').trim();
    
    // 取得自訂連結 (若沒設定，按鈕會導向預設值)
    const sheetViewUrl = (process.env.SHEET_VIEW_URL || '').replace(/"/g, '').trim();
    const shopUrl = (process.env.SHOP_URL || '').replace(/"/g, '').trim();

    if (lineToken && lineIdsString) {
      const lineIds = lineIdsString.split(',');
      const messages = buildFlexOrTextMessages(formData, sheetViewUrl, shopUrl);

      await fetch('https://api.line.me/v2/bot/message/multicast', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${lineToken}`
        },
        body: JSON.stringify({
          to: lineIds,
          messages: messages
        })
      });
    }

    return res.status(200).json({ success: true, message: '資料提交成功' });

  } catch (error) {
    console.error("API 錯誤:", error);
    return res.status(500).json({ success: false, message: error.message });
  }
}