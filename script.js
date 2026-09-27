let count = 0;

// Excel ustunlar xaritasi (B, F, H, J, M, Q, S, W, Y, AA, AD, AG, AJ)
const FIELD_COLUMNS = {
    fish: 2, birth: 6, nationality: 8, gender: 10, father: 13,
    father_job: 17, mother: 19, mother_job: 23, phone: 25,
    address: 27, mahalla: 30, metrika: 33, jshshir: 36
};

// Merge patternlar (ustun indekslari)
const MERGE_PATTERNS = [
    [2, 5], [6, 7], [8, 9], [10, 12], [13, 16], [17, 18],
    [19, 22], [23, 24], [25, 26], [27, 29], [30, 32], [33, 35], [36, 38]
];

const FIRST_DATA_ROW = 2;

// Yangi o'quvchi qo'shish
function addStudent() {
    count++;
    const template = document.getElementById("studentTemplate");
    const clone = template.content.cloneNode(true);
    const card = clone.querySelector(".student");
    card.dataset.id = count;
    card.querySelector(".number").textContent = count;
    
    const birthInput = card.querySelector(".birth-input");
    
    flatpickr(birthInput, {
        dateFormat: "d/m/Y",
        allowInput: true
    });

    birthInput.addEventListener("input", (e) => {
        let val = e.target.value.replace(/\D/g, "");
        if (val.length >= 2 && val.length < 4) {
            val = val.substring(0, 2) + "/" + val.substring(2);
        } else if (val.length >= 4) {
            val = val.substring(0, 2) + "/" + val.substring(2, 4) + "/" + val.substring(4, 8);
        }
        e.target.value = val;
    });

    document.getElementById("students").appendChild(clone);
    renumber();
}

// O'quvchini o'chirish
function removeStudent(button) {
    const cards = document.querySelectorAll(".student");
    if (cards.length === 1) {
        showError("Kamida 1 ta o‘quvchi bo‘lishi kerak.");
        return;
    }
    button.closest(".student").remove();
    renumber();
}

// O'quvchilar tartib raqamini qayta raqamlash
function renumber() {
    document.querySelectorAll(".student").forEach((card, index) => {
        card.querySelector(".number").textContent = index + 1;
    });
}

// Formadagi ma'lumotlarni yig'ish
function collectStudents() {
    return [...document.querySelectorAll(".student")].map(card => {
        const obj = {};
        card.querySelectorAll("[data-key]").forEach(input => {
            obj[input.dataset.key] = input.value.trim();
        });
        return obj;
    });
}

// Xatoliklarni ko'rsatish
function showError(message) {
    const el = document.getElementById("error");
    el.textContent = message;
    setTimeout(() => el.textContent = "", 4000);
}

// Excel yaratish va yuklab berish
async function generateExcel() {
    const className = document.getElementById("className").value.trim();
    const students = collectStudents();

    if (!className) {
        showError("Avval sinf nomini yozing.");
        document.getElementById("className").focus();
        return;
    }

    if (!students.length) {
        showError("Kamida 1 ta o‘quvchi qo‘shing.");
        return;
    }

    for (let i = 0; i < students.length; i++) {
        if (!students[i].fish) {
            showError(`${i + 1}-o‘quvchining F.I.Sh. kiritilmagan.`);
            return;
        }
    }

    const button = document.querySelector(".excel-btn");
    button.disabled = true;
    button.textContent = "EXCEL TAYYORLANMOQDA...";

    try {
        // 1. Root katalogdagi template.xlsx faylini o'qish
        const response = await fetch("template.xlsx");
        if (!response.ok) throw new Error("template.xlsx shablon fayli topilmadi!");
        
        const arrayBuffer = await response.arrayBuffer();
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(arrayBuffer);

        const worksheet = workbook.worksheets[0];

        // 2. 26-qatordan ko'p bo'lsa merge patternlarini davom ettirish
        const lastTemplateRow = 3;
        const neededLastRow = FIRST_DATA_ROW + students.length - 1;

        if (neededLastRow > lastTemplateRow) {
            for (let r = lastTemplateRow + 1; r <= neededLastRow; r++) {
                MERGE_PATTERNS.forEach(([startCol, endCol]) => {
                    worksheet.mergeCells(r, startCol, r, endCol);
                });
            }
        }

        // 3. Ma'lumotlarni yozish
        students.forEach((student, index) => {
            const rowNum = FIRST_DATA_ROW + index;
            const row = worksheet.getRow(rowNum);

            // T/r (№)
            row.getCell(1).value = index + 1;

            // Maydonlar
            Object.keys(FIELD_COLUMNS).forEach(key => {
                const colNum = FIELD_COLUMNS[key];
                let value = student[key] || "";
                const cell = row.getCell(colNum);

                if (key === "birth" && value) {
                    const parts = value.split('/');
                    if (parts.length === 3) {
                        const dateObj = new Date(parts[2], parts[1] - 1, parts[0]);
                        cell.value = dateObj;
                        cell.numFmt = 'dd.mm.yyyy';
                    } else {
                        cell.value = String(value);
                    }
                } else {
                    cell.value = String(value).trim();
                }
            });

            row.commit();
        });

        // 4. Faylni saqlash va yuklash
        const buffer = await workbook.xlsx.writeBuffer();
        const cleanName = className.replace(/[\\/:*?"<>|]+/g, "_") || "sinf";
        const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        
        saveAs(blob, `${cleanName}.xlsx`);

    } catch (err) {
        showError(err.message || "Excel yaratishda xatolik.");
    } finally {
        button.disabled = false;
        button.textContent = "EXCEL YARATISH";
    }
}

// Sayt yuklanganda 1-o'quvchini qo'shish
addStudent();