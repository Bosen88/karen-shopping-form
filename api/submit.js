import { GoogleSpreadsheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';

export default async function handler(req, res) {
  // 只允許 POST 請求
  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, message: 'Method Not Allowed' });
  }

  try {
    const formData = req.body;

    // 1. 初始化 Google Sheets 驗證
    // 將 Vercel 環境變數中的 \n 還原成真實的換行符號
    const privateKey = process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n');
    
    const serviceAccountAuth = new JWT({
      email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      key: privateKey,
      scopes: ['[https://www.googleapis.com/auth/spreadsheets](https://www.googleapis.com/auth/spreadsheets)'],
    });

    const doc = new GoogleSpreadsheet(process.env.GOOGLE_SHEET_ID, serviceAccountAuth);
    await doc.loadInfo(); 
    
    // 取得指定工作表
    const sheet = doc.sheetsByTitle['顧客輪廓分析'];
    if (!sheet) throw new Error('找不到名為 "顧客輪廓分析" 的工作表');

    // 2. 準備要寫入試算表的資料 (對應你的標題列)
    const newRow = {
      'Email': formData.email || "未提供Email",
      '姓名': formData.name || "未提供姓名",
      'LINE暱稱': formData.lineName_placeholder || "",
      '生日-年': formData.birthYear || "",
      '生日-月': formData.birthMonth || "",
      '主要興趣類別': formData.mainCategories ? formData.mainCategories.join(', ') : "",
      '所有勾選項目': formData.allInterests ? formData.allInterests.join(', ') : "",
      // ... 這裡可以加上你需要的解析邏輯 (SkinTypes, etc.)，為了簡潔先用原本字串
      '品牌偏好': formData.brandPref || "",
      '回購頻率': formData.repurchaseFrequency || "",
      '是否願意折扣通知': formData.consentDiscount || "",
      '最在意的三項': formData.top3Needs || "",
      '最後更新時間': new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' })
    };

    // 寫入 Google Sheet
    await sheet.addRow(newRow);

    // 3. 發送 LINE 通知 (如果有設定 Token 的話)
    if (process.env.LINE_CHANNEL_ACCESS_TOKEN && process.env.LINE_RECIPIENT_USER_IDS) {
      const lineIds = process.env.LINE_RECIPIENT_USER_IDS.split(',');
      const lineMessage = {
        to: lineIds,
        messages: [{
          type: "text",
          text: `🎉 新回覆：${formData.name}\nEmail：${formData.email}\n優先：${formData.top3Needs}`
        }]
      };

      await fetch('[https://api.line.me/v2/bot/message/multicast](https://api.line.me/v2/bot/message/multicast)', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${process.env.LINE_CHANNEL_ACCESS_TOKEN}`
        },
        body: JSON.stringify(lineMessage)
      });
    }

    // 回傳成功給前端
    return res.status(200).json({ success: true, message: '資料提交成功' });

  } catch (error) {
    console.error("API 錯誤:", error);
    return res.status(500).json({ success: false, message: error.message });
  }
}
