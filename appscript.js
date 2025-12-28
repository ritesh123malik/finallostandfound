
const SHEET_NAME = "Items";
const USERS_SHEET_NAME = "Users"; 
const ADMIN_EMAIL = "riteshmalik21092005@gmail.com"; 

const SPREADSHEET_ID = "1dz3omCfayBlEX5AIqJh10uyEDcOtIBei56aI8vTtdMg"; 


function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);

    const data = sheet.getDataRange().getValues();
    if (data.length > 1) data.shift(); else return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);

    const json = data.map((row, index) => ({ 
      rowIndex: index + 2, id: row[0], type: row[1], item: row[2], desc: row[3], status: row[4], reporter: row[5], email: row[6], date: row[7], image: row[8] 
    }));
    return ContentService.createTextOutput(JSON.stringify(json)).setMimeType(ContentService.MimeType.JSON);
  } catch(e) { return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON); }
}


function doPost(e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const params = JSON.parse(e.postData.contents);
    
   
    if (params.action === "REPORT") {
      let sheet = ss.getSheetByName(SHEET_NAME);
      if(!sheet) { sheet = ss.insertSheet(SHEET_NAME); sheet.appendRow(["ID", "Type", "Item", "Desc", "Status", "Reporter", "Email", "Date", "Image"]); }
      
      const id = new Date().getTime().toString();
      sheet.appendRow([id, params.type, params.item, params.desc, "Open", params.reporter, params.email, new Date(), params.image || ""]);
      
      try {
        const userSheet = ss.getSheetByName(USERS_SHEET_NAME);
        if(userSheet) {
           const recipients = [...new Set(userSheet.getDataRange().getValues().slice(1).map(r => r[0]).filter(e => e && e.toString().includes("@")))];
           
           if(recipients.length > 0) {
             // Define colors based on status
             const isLost = params.type === 'Lost';
             const badgeColor = isLost ? '#fee2e2' : '#dcfce7'; // Red bg vs Green bg
             const textColor = isLost ? '#ef4444' : '#16a34a'; // Red text vs Green text
             const badgeText = isLost ? '🔴 LOST ITEM' : '🟢 FOUND ITEM';
             
             MailApp.sendEmail({
               to: ADMIN_EMAIL,
               bcc: recipients.join(","),
               subject: `${isLost ? '🔴' : '🟢'} Alert: ${params.item} reported as ${params.type}`,
               htmlBody: `
                 <div style="font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 16px; overflow: hidden; box-shadow: 0 10px 40px rgba(0,0,0,0.1); border: 1px solid #f0f0f0;">
                    
                    <div style="background: linear-gradient(135deg, #4f46e5 0%, #7c3aed 100%); padding: 40px 20px; text-align: center;">
                        <h1 style="color: #ffffff; margin: 0; font-size: 28px; font-weight: 800; letter-spacing: -0.5px;">Lost<span style="font-weight: 300; opacity: 0.8;">&</span>Found</h1>
                        <p style="color: rgba(255,255,255,0.9); margin: 8px 0 0; font-size: 13px; text-transform: uppercase; letter-spacing: 2px;">Campus Recovery Network</p>
                    </div>

                    <div style="padding: 40px 30px;">
                        
                        <div style="text-align: center; margin-bottom: 25px;">
                            <span style="background-color: ${badgeColor}; color: ${textColor}; padding: 8px 20px; border-radius: 100px; font-size: 12px; font-weight: 800; letter-spacing: 1px; display: inline-block;">
                                ${badgeText}
                            </span>
                        </div>

                        <h2 style="text-align: center; color: #1e293b; font-size: 32px; margin: 0 0 30px 0; font-weight: 800; line-height: 1.2;">${params.item}</h2>

                        ${params.image ? `
                        <div style="margin-bottom: 30px; text-align: center;">
                            <img src="${params.image}" style="width: 100%; max-width: 100%; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.08); border: 1px solid #e2e8f0; object-fit: cover;">
                        </div>` : ''}

                        <div style="background-color: #f8fafc; border-radius: 16px; padding: 25px; border: 1px solid #e2e8f0;">
                            <div style="margin-bottom: 15px;">
                                <p style="margin: 0; color: #64748b; font-size: 11px; text-transform: uppercase; font-weight: 700; letter-spacing: 0.5px;">Description</p>
                                <p style="margin: 5px 0 0; color: #334155; font-size: 15px; line-height: 1.6;">${params.desc.split("||")[0]}</p>
                            </div>
                            
                            <div style="border-top: 1px solid #e2e8f0; margin: 15px 0;"></div>

                            <div style="display: flex; justify-content: space-between;">
                                <div>
                                    <p style="margin: 0; color: #64748b; font-size: 11px; text-transform: uppercase; font-weight: 700; letter-spacing: 0.5px;">Contact</p>
                                    <p style="margin: 5px 0 0; color: #334155; font-size: 14px;">${params.email}</p>
                                </div>
                                <div style="text-align: right;">
                                    <p style="margin: 0; color: #64748b; font-size: 11px; text-transform: uppercase; font-weight: 700; letter-spacing: 0.5px;">Date</p>
                                    <p style="margin: 5px 0 0; color: #334155; font-size: 14px;">${new Date().toLocaleDateString()}</p>
                                </div>
                            </div>
                        </div>

                        <div style="text-align: center; margin-top: 35px;">
                            <a href="#" style="background: linear-gradient(135deg, #4f46e5 0%, #7c3aed 100%); color: #ffffff; padding: 16px 32px; text-decoration: none; border-radius: 100px; font-weight: 700; font-size: 14px; display: inline-block; box-shadow: 0 10px 20px -10px rgba(79, 70, 229, 0.5);">
                                Check Portal
                            </a>
                        </div>
                    </div>

                    <div style="background-color: #f1f5f9; padding: 25px; text-align: center; border-top: 1px solid #e2e8f0;">
                        <p style="margin: 0; color: #94a3b8; font-size: 11px;">© 2025 College Gymkhana. Automated Notification.</p>
                    </div>
                 </div>
               `
             });
           }
        }
      } catch(e) {}

      return response({ result: "Success" });
    } 
    
    
    else if (params.action === "REGISTER_USER") {
      let userSheet = ss.getSheetByName(USERS_SHEET_NAME);
      if (!userSheet) { userSheet = ss.insertSheet(USERS_SHEET_NAME); userSheet.appendRow(["Email", "Date"]); }
      const data = userSheet.getDataRange().getValues();
      if (!data.some(r => r[0] === params.email)) userSheet.appendRow([params.email, new Date()]);
      return response({ result: "Registered" });
    }
    else if (params.action === "RESOLVE") {
      const sheet = ss.getSheetByName(SHEET_NAME);
      sheet.getRange(params.rowIndex, 5).setValue("Resolved");
      return response({ result: "Resolved" });
    }
    else if (params.action === "BROADCAST") {
      const userSheet = ss.getSheetByName(USERS_SHEET_NAME);
      if(!userSheet) return response({ result: "Error" });
      const recipients = [...new Set(userSheet.getDataRange().getValues().slice(1).map(r => r[0]).filter(e => e && e.toString().includes("@")))];
      if(recipients.length > 0) {
        MailApp.sendEmail({ to: ADMIN_EMAIL, bcc: recipients.join(","), subject: `📢 ${params.subject}`, htmlBody: `<h3>${params.message}</h3>` });
        return response({ result: "Success" });
      }
      return response({ result: "No Users" });
    }
    
  } catch (error) { return response({ result: "Error", message: error.toString() }); }
}

function response(data) { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }
