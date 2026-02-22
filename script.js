// =====================
// بيانات الملفات
// =====================
let singleFileDataGlobal = {}; // للملف الواحد
let file1DataGlobal = {}; // للملف الأول عند الملفين
let file2DataGlobal = {}; // للملف الثاني عند الملفين

// =====================
// دوال مساعدة
// =====================
function normalizeArabic(text) {
    return text.toString().replace(/ـ/g, "").replace(/\s+/g, "").trim();
}

function getMostFrequent(arr) {
    let counts = {};
    arr.forEach(val => { if (!val) return; counts[val] = (counts[val] || 0) + 1; });
    return Object.entries(counts).sort((a,b)=>b[1]-a[1])[0]?.[0]||"-";
}

function createGoogleMapsLink(coord){

    if(!coord) return "-";

    let value = coord.toString().trim();

    // إزالة المسافات الزائدة
    value = value.replace(/\s+/g, " ");

    // إذا كانت قيم غير صالحة
    if(
        value === "" ||
        value === "-" ||
        value === "." ||
        value === "," ||
        value === "،"
    ){
        return value;
    }

    // 🔥 تحقق أن القيمة إحداثيات صحيحة (رقم, رقم)
    const coordRegex = /^-?\d+(\.\d+)?\s*,\s*-?\d+(\.\d+)?$/;

    if(!coordRegex.test(value)){
        return value; // ليست إحداثيات → لا تحولها لرابط
    }

    // إنشاء رابط Google Maps
    let encoded = encodeURIComponent(value);

    return `<a href="https://www.google.com/maps/search/?api=1&query=${encoded}" 
               target="_blank" 
               style="color:#1f4e79;text-decoration:underline;">
               ${value}
            </a>`;
}


// =====================
// تعديل أسماء الملفات لعرض واضح بدون امتداد
// =====================
function getFileNameWithoutExtension(filename){
    if(!filename) return "";
    return filename.replace(/\.[^/.]+$/, ""); // يشيل آخر امتداد
}


// =====================
// تحويل تاريخ Excel ل JS Date
// =====================
function excelDateToJSDate(serial) {
    if(!serial) return "";
    const utc_days  = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;                                        
    const date_info = new Date(utc_value * 1000);
    const fractional_day = serial - Math.floor(serial) + 0.0000001;
    let total_seconds = Math.floor(86400 * fractional_day);
    const seconds = total_seconds % 60;
    total_seconds -= seconds;
    const hours = Math.floor(total_seconds / (60*60));
    const minutes = Math.floor(total_seconds / 60) % 60;
    return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds);
}

function formatDate(date){
    if(!date) return "";
    const dd = String(date.getDate()).padStart(2,'0');
    const mm = String(date.getMonth()+1).padStart(2,'0');
    const yyyy = date.getFullYear();
    return `${dd}/${mm}/${yyyy}`;
}

// =====================
// توليد جدول للسجلات
// =====================
function generateTableForRows(rows) {
    if (!rows || !rows.length) return "<p>لا توجد بيانات</p>";

    let html = `<div style="max-height:300px;overflow:auto;border:1px solid #ccc;margin-bottom:20px;">
        <table>
            <tr>
                <th>الاسم</th>
                <th>الإحداثيات</th>
                <th>نوع المكالمة</th>
                <th>المدة</th>
                <th>وقت الاتصال</th>
                <th>تاريخ الاتصال</th>
            </tr>`;

    rows.forEach(r=>{
        let coords = createGoogleMapsLink(r["الإحداثيات"]);
        html += `<tr>
            <td>${r["الاسم"]||"-"}</td>
            <td style="direction:ltr;">${coords}</td>
            <td>${r["نوع المكالمة"]||"-"}</td>
            <td>${r["المدة"]||"-"}</td>
            <td>${r["وقت الاتصال"]||"-"}</td>
            <td>${r["تاريخ الاتصال"]||"-"}</td>
        </tr>`;
    });

    html += `</table></div>`;
    return html;
}

// =====================
// قراءة أي ملف
// =====================
function readExcelAndStoreData(file, targetData) {
    return new Promise((resolve,reject)=>{
        const reader = new FileReader();
        reader.onload = function(e){
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data,{type:"array"});
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(sheet,{header:1,defval:""});

            let headerRow=-1;
            let serviceCol=-1;

            // 🔥 نبحث عن صف العناوين الخاص بجدول المكالمات فقط
            for(let r=0;r<rows.length;r++){
                for(let c=0;c<rows[r].length;c++){
                    if(!rows[r][c]) continue;
                    let cell = normalizeArabic(rows[r][c]);
                    if(cell.includes("رقمالخدمة")){
                        headerRow=r;
                        serviceCol=c;
                        break;
                    }
                }
                if(headerRow!==-1) break;
            }

            if(headerRow===-1){
                alert("لم يتم العثور على جدول سجل المكالمات");
                resolve([]);
                return;
            }

            const headers = rows[headerRow];

            // تحديد الأعمدة من نفس صف الهيدر
            let nameCol=-1, durationCol=-1, callTypeCol=-1,
                callTimeCol=-1, callDateCol=-1, coordinatesCol=-1;

            headers.forEach((h,i)=>{
                let key = normalizeArabic(h);
                if(key.includes("الاسم")) nameCol=i;
                if(key.includes("المدة")) durationCol=i;
                if(key.includes("نوعالمكالمة")) callTypeCol=i;
                if(key.includes("وقتالاتصال")) callTimeCol=i;
                if(key.includes("تاريخالاتصال")) callDateCol=i;
                if(key.includes("الإحداثيات")) coordinatesCol=i;
            });

            let serviceNumbers=[];

            // 🔥 نقرأ فقط الصفوف بعد صف العناوين
            for(let r=headerRow+1;r<rows.length;r++){

                let serviceValue = rows[r][serviceCol];
                if(!serviceValue) continue;

                serviceValue = serviceValue.toString().trim();
                serviceNumbers.push(serviceValue);

                if(!targetData[serviceValue])
                    targetData[serviceValue]=[];

                let rowObject = {
                    "الاسم": nameCol!==-1 ? rows[r][nameCol] || "" : "",
                    "المدة": durationCol!==-1 ? rows[r][durationCol] || "" : "",
                    "نوع المكالمة": callTypeCol!==-1 ? rows[r][callTypeCol] || "" : "",
                    "وقت الاتصال": callTimeCol!==-1 ? rows[r][callTimeCol] || "" : "",
                    "الإحداثيات": coordinatesCol!==-1 ? rows[r][coordinatesCol] || "" : "",
                    "تاريخ الاتصال": ""
                };

                if(callDateCol!==-1){
                    let dateCell = rows[r][callDateCol];
                    if(typeof dateCell==="number")
                        rowObject["تاريخ الاتصال"] =
                            formatDate(excelDateToJSDate(dateCell));
                    else
                        rowObject["تاريخ الاتصال"] = dateCell || "";
                }

                targetData[serviceValue].push(rowObject);
            }

            resolve(serviceNumbers);
        };
        reader.readAsArrayBuffer(file);
    });
}

// =====================
// تحليل ملف واحد
// =====================
let currentOpen = null;
async function analyzeSingle(){
    const file=document.getElementById("file1").files[0];
    if(!file){ alert("اختر ملف أولاً"); return; }

    // تفريغ الملف الثاني تلقائياً
    document.getElementById("file2").value = "";

    singleFileDataGlobal = {};
    const numbers = await readExcelAndStoreData(file,singleFileDataGlobal);

    let counts={};
    numbers.forEach(num=>{ counts[num]=(counts[num]||0)+1; });

    const sorted = Object.entries(counts).sort((a,b)=>b[1]-a[1]);
    displayResultsSingle(sorted);
}

function displayResultsSingle(data){
  
    const resultDiv = document.getElementById("result");
    let html=`<table><tr><th>رقم الخدمة</th><th>عدد التكرار</th></tr>`;
    data.forEach(item=>{
        html += `<tr>
            <td class="number single" onclick="showDetailsSingle('${item[0]}')">
                ${item[0]}
            </td>
            <td>${item[1]}</td>
        </tr>`;
    });
    html += `</table><div id="detailsContainer" style="margin-top:30px;"></div>`;
    resultDiv.innerHTML = html;
}

function showDetailsSingle(serviceNumber){
    const container = document.getElementById("detailsContainer");
    if(currentOpen===serviceNumber){ container.innerHTML=""; currentOpen=null; return; }
    currentOpen=serviceNumber;

    const rows = singleFileDataGlobal[serviceNumber] || [];
    if(!rows.length) return;

    const names=rows.map(r=>r["الاسم"]||"-");
    const locations = rows.map(r => r["الإحداثيات"] || "-");
    const callTypes=rows.map(r=>r["نوع المكالمة"]||"-");
    const mostName=getMostFrequent(names);
    const mostLocation=createGoogleMapsLink(getMostFrequent(locations));
    const mostCallType=getMostFrequent(callTypes);

    let html=`<h3>تفاصيل رقم الخدمة: ${serviceNumber}</h3>
        <table>
            <tr><th>الاسم الأكثر تكرار</th><td>${mostName}</td></tr>
            <tr><th>الإحداثيات الأكثر تكرار</th><td style="direction:ltr;">${mostLocation}</td></tr>
            <tr><th>نوع المكالمة الأكثر تكرار</th><td>${mostCallType}</td></tr>
            <tr><th>إجمالي عدد السجلات</th><td>${rows.length}</td></tr>
        </table>
        <h3 style="margin-top:20px">جميع السجلات</h3>
        ${generateTableForRows(rows)}`;
    container.innerHTML=html;
}

// =====================
// تحليل ملفين
// =====================
let currentCommonOpen = null;

async function analyzeBothFiles(){
    const file1=document.getElementById("file1").files[0];
    const file2=document.getElementById("file2").files[0];
    if(!file1||!file2){ alert("اختر كلا الملفين أولاً"); return; }

    file1Name = getFileNameWithoutExtension(file1.name);
    file2Name = getFileNameWithoutExtension(file2.name);

    file1DataGlobal={}; 
    file2DataGlobal={};

    const numbers1 = await readExcelAndStoreData(file1,file1DataGlobal);
    const numbers2 = await readExcelAndStoreData(file2,file2DataGlobal);

    let counts={};

    numbers1.forEach(num=>{
        counts[num]=(counts[num]||0)+1;
    });

    numbers2.forEach(num=>{
        counts[num]=(counts[num]||0)+1;
    });

    const commonNumbers = Object.keys(counts).filter(num =>
        numbers1.includes(num) && numbers2.includes(num)
    );

    if(!commonNumbers.length){
        alert("لا يوجد أرقام مشتركة بين الملفين");
        document.getElementById("result").innerHTML="";
        return;
    }

    // 🔥 الترتيب من الأعلى تكرار إلى الأقل
    commonNumbers.sort((a,b)=>counts[b] - counts[a]);

    const resultDiv=document.getElementById("result");
    let html=`<table>
        <tr>
            <th>رقم الخدمة المشترك</th>
            <th>عدد التكرار</th>
        </tr>`;

    commonNumbers.forEach(num=>{
        html += `<tr>
            <td class="number common" onclick="showCommonDetails('${num}')">
                ${num}
            </td>
            <td>${counts[num]}</td>
        </tr>`;
    });

    html += `</table>
        <div id="detailsContainer" style="margin-top:30px;"></div>`;

    resultDiv.innerHTML=html;
}


function showCommonDetails(serviceNumber){
    const container=document.getElementById("detailsContainer");

    if(currentCommonOpen===serviceNumber){
        container.innerHTML="";
        currentCommonOpen=null;
        return;
    }

    currentCommonOpen=serviceNumber;

    const rows1=file1DataGlobal[serviceNumber]||[];
    const rows2=file2DataGlobal[serviceNumber]||[];

    const allRows = [...rows1, ...rows2];
    const locations = allRows.map(r => r["الإحداثيات"] || "-");
    const mostLocation = createGoogleMapsLink(getMostFrequent(locations));

    let html=`
        <h3>تفاصيل رقم الخدمة المشترك: ${serviceNumber}</h3>

        <table>
            <tr>
                <th>الإحداثيات الأكثر تكرار</th>
                <td style="direction:ltr;">${mostLocation}</td>
            </tr>
        </table>

        <h4>الملف الأول: ${file1Name}</h4>
        ${generateTableForRows(rows1)}

        <h4>الملف الثاني: ${file2Name}</h4>
        ${generateTableForRows(rows2)}
    `;

    container.innerHTML=html;
}