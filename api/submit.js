import { GoogleSpreadsheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';

export default async function handler(req, res) {
  // 只允許 POST 請求
  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, message: 'Method Not Allowed' });
  }

  try {
    const formData = req.body;

    // ==========================================
    // 🛡️ 終極防呆機制：自動清理 Vercel 環境變數的雜訊
    // ==========================================
    
    // 1. 清理 Email (去除可能不小心貼上的雙引號與前後空白)
    const clientEmail = (process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || '').replace(/"/g, '').trim();
    
    // 2. 清理 Private Key (移除雙引號、確保換行符號正確解析)
    let privateKey = process.env.GOOGLE_PRIVATE_KEY || '';
    privateKey = privateKey.replace(/"/g, '') // 移除所有雙引號
                           .replace(/\\n/g, '\n') // 將字串的 \n 轉為真實的斷行
                           .trim(); // 移除前後多餘空白

    // 3. 清理 Sheet ID
    const sheetId = (process.env.GOOGLE_SHEET_ID || '').replace(/"/g, '').trim();

    // 檢查是否有抓到變數，避免出現 undefined 錯誤
    if (!clientEmail || !privateKey || !sheetId) {
        throw new Error('環境變數遺失：請檢查 Vercel 後台是否已設定 GOOGLE_SERVICE_ACCOUNT_EMAIL, GOOGLE_PRIVATE_KEY, GOOGLE_SHEET_ID');
    }

    // ==========================================

    // 初始化 Google Sheets 驗證
    const serviceAccountAuth = new JWT({
      email: clientEmail,
      key: privateKey,
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const doc = new GoogleSpreadsheet(sheetId, serviceAccountAuth);
    
    // 這裡如果成功，代表金鑰完美運作！
    await doc.loadInfo(); 
    
    // 取得指定工作表
    const sheet = doc.sheetsByTitle['顧客輪廓分析'];
    if (!sheet) throw new Error('找不到名為 "顧客輪廓分析" 的工作表，請檢查 Google 試算表下方的頁籤名稱');

    // 準備要寫入試算表的資料 (對應你的標題列)
    const newRow = {
      'Email': formData.email || "未提供Email",
      '姓名': formData.name || "未提供姓名",
      'LINE暱稱': formData.lineName_placeholder || "",
      '生日-年': formData.birthYear || "",
      '生日-月': formData.birthMonth || "",
      '主要興趣類別': formData.mainCategories ? formData.mainCategories.join(', ') : "",
      '所有勾選項目': formData.allInterests ? formData.allInterests.join(', ') : "",
      '品牌偏好': formData.brandPref || "",
      '回購頻率': formData.repurchaseFrequency || "",
      '是否願意折扣通知': formData.consentDiscount || "",
      '最在意的三項': formData.top3Needs || "",
      '最後更新時間': new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' })
    };

    // 寫入 Google Sheet
    await sheet.addRow(newRow);

    // 發送 LINE 通知 (如果有設定 Token，一樣加入防呆清理)
    const lineToken = (process.env.LINE_CHANNEL_ACCESS_TOKEN || '').replace(/"/g, '').trim();
    const lineIdsString = (process.env.LINE_RECIPIENT_USER_IDS || '').replace(/"/g, '').trim();

    if (lineToken && lineIdsString) {
      const lineIds = lineIdsString.split(',');
      const lineMessage = {
        to: lineIds,
        messages: [{
          type: "text",
          text: `🎉 新回覆：${formData.name}\nEmail：${formData.email}\n優先：${formData.top3Needs}`
        }]
      };

      await fetch('https://api.line.me/v2/bot/message/multicast', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${lineToken}`
        },
        body: JSON.stringify(lineMessage)
      });
    }

    // 回傳成功給前端
    return res.status(200).json({ success: true, message: '資料提交成功' });

  } catch (error) {
    console.error("API 錯誤:", error);
    // 把詳細的錯誤訊息傳回前端，方便除錯
    return res.status(500).json({ success: false, message: error.message });
  }
}