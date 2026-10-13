let emailData = {};

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("processBtn").onclick = processExcel;
        
        // استخراج بيانات الإيميل أوتوماتيكياً
        const item = Office.context.mailbox.item;
        
        emailData.subject = item.subject || "بدون موضوع";
        emailData.senderName = item.sender.displayName || item.sender.emailAddress;
        emailData.date = item.dateTimeCreated;
        
        // تصنيف تلقائي بناءً على الكلمات المفتاحية في الموضوع
        let nature = "Demande ";
        let subjLower = emailData.subject.toLowerCase();
        if(subjLower.includes("reclamation") || subjLower.includes("réclamation") || subjLower.includes("شكوى")) {
            nature = "Réclamation ";
        }
        emailData.nature = nature;

        // عرض المعلومات للمستخدم للتأكيد
        document.getElementById("emailInfo").innerHTML = 
            "<b>طبيعة الحدث:</b> " + emailData.nature + "<br/>" +
            "<b>المُرسل:</b> " + emailData.senderName + "<br/>" +
            "<b>الموضوع:</b> " + emailData.subject;
    }
});

function processExcel() {
    const fileInput = document.getElementById("fileInput");
    const statusDiv = document.getElementById("status");

    if (!fileInput.files || fileInput.files.length === 0) {
        statusDiv.innerHTML = "<span style='color: red;'>الرجاء اختيار ملف Excel أولاً.</span>";
        return;
    }

    statusDiv.innerText = "جاري معالجة الملف...";
    const file = fileInput.files[0];
    const reader = new FileReader();

    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            // قراءة الملف مع الحفاظ على التنسيقات والخلايا
            const workbook = XLSX.read(data, { type: 'array', cellDates: true, cellStyles: true });
            
            const sheetName = "S1";
            if (!workbook.Sheets[sheetName]) {
                statusDiv.innerHTML = "<span style='color: red;'>خطأ: الورقة 'S1' غير موجودة في الملف.</span>";
                return;
            }
            
            let sheet = workbook.Sheets[sheetName];
            
            // تحويل ورقة الإكسيل إلى مصفوفة (JSON) لنتمكن من قراءتها
            // في ملفك، العناوين الحقيقية توجد في الصف 22، لذلك نتجاهل ما قبله
            let rows = XLSX.utils.sheet_to_json(sheet, { range: 22, raw: false });
            
            // تحديد آخر رقم تسلسلي N°
            let lastId = 0;
            if(rows.length > 0) {
                 const lastRow = rows[rows.length - 1];
                 // البحث عن العمود المسمى N° (مع أخذ المسافات المحتملة بعين الاعتبار)
                 const idKey = Object.keys(lastRow).find(k => k.trim() === 'N°');
                 if(idKey && !isNaN(lastRow[idKey])) {
                     lastId = parseInt(lastRow[idKey]);
                 }
            }

            // تنسيق التاريخ والوقت
            const d = new Date(emailData.date);
            const dateStr = d.toLocaleDateString('en-GB'); // تنسيق DD/MM/YYYY
            const timeStr = d.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});

            // بناء السطر الجديد الذي سيضاف للإكسيل
            let newRow = {};
            newRow["N°"] = lastId + 1;
            newRow["Date"] = dateStr;
            newRow["Heure"] = timeStr;
            newRow["Nature "] = emailData.nature;
            newRow["Service"] = "Transport"; // قيمة افتراضية يمكن تعديلها في الإكسيل لاحقاً
            newRow["Désignation "] = emailData.subject; // وضع موضوع الإيميل في خانة Désignation
            newRow["Déclarant"] = emailData.senderName; // اسم المرسل
            newRow["Responsable du traitement "] = "ADEL KHEDIMI";
            newRow["Status"] = "Nouveau";
            
            // إضافة السطر للمصفوفة
            rows.push(newRow);

            // إنشاء ورقة جديدة تحتوي على البيانات المُحدثة بدءاً من الصف 23
            let newSheet = XLSX.utils.json_to_sheet(rows, {origin: "A23"});
            
            // الحفاظ على الترويسة القديمة (الصفوف 1 إلى 22)
            // هذه الخطوة تضمن عدم مسح الديكور في أعلى الملف
            const oldSheet = workbook.Sheets[sheetName];
            for (let cell in oldSheet) {
                if (cell[0] === '!') continue; // تخطي معلومات التنسيق
                
                // استخراج رقم الصف
                let rowNum = parseInt(cell.replace(/[^0-9]/g, ''));
                if (rowNum <= 22) {
                    newSheet[cell] = oldSheet[cell];
                }
            }
            
            // الحفاظ على نطاق الورقة
            newSheet['!ref'] = XLSX.utils.encode_range({
                s: { c: 0, r: 0 },
                e: { c: 15, r: 22 + rows.length } // 15 عموداً، وعدد الصفوف الجديد
            });

            // استبدال الورقة القديمة بالجديدة
            workbook.Sheets[sheetName] = newSheet;

            // توليد اسم جديد للملف المحمل
            const today = new Date();
            const dateSuffix = today.getDate() + "-" + (today.getMonth()+1) + "-" + today.getFullYear();
            const newFileName = "Mis_a_jour_" + dateSuffix + "_" + file.name;

            // تحميل الملف
            XLSX.writeFile(workbook, newFileName);
            
            statusDiv.innerHTML = "<span style='color: green;'>✅ تم التصدير بنجاح! تم تحميل الملف المحدث.</span>";
        } catch (error) {
            console.error(error);
            statusDiv.innerHTML = "<span style='color: red;'>❌ حدث خطأ غير متوقع.</span>";
        }
    };

    reader.readAsArrayBuffer(file);
}