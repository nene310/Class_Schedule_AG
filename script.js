/**
 * ScheduleAG - Core Logic
 */

// Global State
let workbook = null;
let rawScheduleData = [];
let defaultTimeSlots = [
    { start: '08:20', end: '09:05' }, // 1
    { start: '09:15', end: '10:00' }, // 2
    { start: '10:20', end: '11:05' }, // 3
    { start: '11:15', end: '12:00' }, // 4
    { start: '14:30', end: '15:15' }, // 5
    { start: '15:25', end: '16:10' }, // 6
    { start: '16:30', end: '17:15' }, // 7
    { start: '17:15', end: '18:00' }, // 8
    { start: '19:10', end: '19:55' }, // 9
    { start: '19:55', end: '20:40' }  // 10
];

// Initialization
document.addEventListener('DOMContentLoaded', () => {
    initTimeSettings();
    document.getElementById('fileUpload').addEventListener('change', handleFileUpload);
    document.getElementById('btnGenerate').addEventListener('click', generateSchedule);
    // document.getElementById('btnPrint').addEventListener('click', () => window.print());
    // document.getElementById('btnExport').addEventListener('click', exportToICS);
});

function initTimeSettings() {
    const container = document.getElementById('timeSettings');
    defaultTimeSlots.forEach((slot, index) => {
        const row = document.createElement('div');
        row.style.display = 'contents';
        row.innerHTML = `
            <span>${index + 1}</span>
            <input type="time" value="${slot.start}" data-idx="${index}" data-type="start">
            <input type="time" value="${slot.end}" data-idx="${index}" data-type="end">
        `;
        container.appendChild(row);
    });
}

// File Handling
// File Handling
function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    document.getElementById('fileName').textContent = "正在读取: " + file.name;

    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);

            if (typeof XLSX === 'undefined') {
                throw new Error("XLSX 库未加载，请检查网络或刷新页面");
            }

            workbook = XLSX.read(data, { type: 'array' });

            if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
                throw new Error("文件解析失败或无工作表");
            }

            // Assume first sheet
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];

            // Convert to JSON (Array of Arrays) to easier handling of messy headers
            rawScheduleData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            if (!rawScheduleData || rawScheduleData.length === 0) {
                throw new Error("工作表为空");
            }

            document.getElementById('fileName').textContent = "已加载: " + file.name;
            console.log("File loaded. Rows:", rawScheduleData.length);

        } catch (err) {
            console.error(err);
            document.getElementById('fileName').textContent = "读取失败: " + err.message;
            alert("读取文件出错: " + err.message);
            rawScheduleData = []; // Reset on error
        }
    };
    reader.readAsArrayBuffer(file);
}

// Core Parsing Logic
function parseCourseString(cellContent) {
    if (!cellContent || typeof cellContent !== 'string') return [];

    // Split by multiple lines if any, but the prompt says single string often just split by regex or just handle the whole block.
    // The prompt says "Cell content may submit multiple courses, separated by newline".
    // Example: "Course1...\nCourse2..."
    const independentCourses = cellContent.split(/\r?\n/).filter(s => s.trim().length > 0);
    const parsedCourses = [];

    independentCourses.forEach(courseStr => {
        // Smart Parsing: Handle variable formats
        // Format A: Name/Code/Weeks/Location/... (Standard)
        // Format B: Name/Weeks/Location (Simplified)
        // Format C: Missing slashes? (Not handled yet, assuming at least some delimiters)

        const parts = courseStr.split('/').map(s => s.trim());

        let name = parts[0];
        let weeks = [];
        let location = "";
        let className = "";

        // Strategy: Find "Week" part specifically
        // It usually contains digit + "周"
        const weekPartIdx = parts.findIndex(p => /(\d+[-~]\d+|\d+)周/.test(p));

        if (weekPartIdx !== -1) {
            // Found Weeks
            weeks = parseWeekString(parts[weekPartIdx]);

            // Name is usually index 0. If week is index 0 (unlikely), name is missing.
            if (weekPartIdx === 0) name = "未知课程";

            // Location is usually AFTER weeks
            // Check parts after weekPartIdx
            for (let i = weekPartIdx + 1; i < parts.length; i++) {
                const p = parts[i];
                // Heuristic: Class name often has "班" or "级" or "专业"
                if (/班|级|专业/.test(p)) {
                    className = p;
                } else if (!location) {
                    // First non-class string after weeks is likely Location
                    // Ignore short codes if they look like nonsense? 
                    location = p;
                }
            }
        } else {
            // Fallback: strict index if no "X周" found (maybe just 1-16 without '周'?)
            if (parts.length >= 4) {
                weeks = parseWeekString(parts[2]); // Try index 2
                location = parts[3];
                className = parts[5] || "";
            } else {
                // If really can't find weeks, we can't schedule it.
                // Try logging error?
                console.warn("Could not parse course string (no weeks found):", courseStr);
                return;
            }
        }

        // 2. Location Simplification
        location = simplifyLocation(location);

        // 3. Name Simplification
        const displayName = simplifyName(name);

        parsedCourses.push({
            rawName: name,
            displayName: displayName,
            weeks: weeks,
            location: location,
            className: className,
            rawStr: courseStr
        });
    });

    return parsedCourses;
}

function parseWeekString(str) {
    // Example: "(1-2节)2-6周,8-12周(双)"
    // Or just "2-6周"
    // We need to extract the Week ranges. The period info (1-2节) might be at start.

    // First, remove the Period info if exists (e.g. (1-2节)) because we know the row index determines the period.
    // However, sometimes period info is useful if one cell covers multiple periods? 
    // The prompt says Excel headers are Weekday(Col) and Period(Row). 
    // Usually the cell at Row X Col Y implies the period. But the string contains "(1-2节)". 
    // We will trust the string for specific period overrides if we were doing advanced mapping, 
    // but for now let's assume the cell position dictates the period, we just scrape weeks.

    // Remove anything inside parens that looks like period "1-2节" to avoid confusion? 
    // Actually the week string is like "2-6周,8-12周(双)".
    // Let's first strip the leading (...) if it contains '节'.
    let cleanStr = str.replace(/\([^)]*节\)/g, "");

    // Logic: Split by comma
    const parts = cleanStr.split(',');
    let weekSet = new Set();

    parts.forEach(part => {
        // Match patterns: "2-6周", "8-12周(双)", "5周"
        // Regex: (\d+)(?:-(\d+))?周(?:\((单|双)\))?
        const match = /(\d+)(?:-(\d+))?周(?:\((单|双)\))?/.exec(part);
        if (match) {
            const start = parseInt(match[1]);
            const end = match[2] ? parseInt(match[2]) : start;
            const type = match[3]; // "单" or "双" or undefined

            for (let i = start; i <= end; i++) {
                if (type === '单' && i % 2 === 0) continue;
                if (type === '双' && i % 2 !== 0) continue;
                weekSet.add(i);
            }
        }
    });

    return Array.from(weekSet).sort((a, b) => a - b);
}

function simplifyLocation(loc) {
    if (!loc) return "";
    let s = loc.replace(/实验实训中心/g, "实训楼");
    // Remove "桂林洋 " or generic campus names if known?
    // Prompt: Remove "冗余校区名", keep "一教N104". 
    // Example: "桂林洋 一教N104" -> "一教N104"
    // A simple heuristic: take the last meaningful part if separated by space?
    // Or just replace known campus names.
    s = s.replace(/桂林洋/g, "").trim();
    return s;
}

function simplifyName(name) {
    // User requested full name display
    return name;
}

// Global Events Store
let generatedEvents = [];
let currentCalendarDate = new Date(); // To track which month we are viewing

function generateSchedule() {
    try {
        console.log("Starting generation...");
        if (rawScheduleData.length === 0) {
            alert("请先上传课表文件");
            return;
        }

        const startDateInput = document.getElementById('semesterStart').value;
        if (!startDateInput) return;
        const semesterStart = new Date(startDateInput);

        // Find Header Row
        let headerRowIdx = -1;
        for (let r = 0; r < rawScheduleData.length; r++) {
            const row = rawScheduleData[r];
            if (row.some(c => c && typeof c === 'string' && c.includes('星期一'))) {
                headerRowIdx = r;
                break;
            }
        }

        if (headerRowIdx === -1) {
            alert("未识别到'星期一'表头，请检查文件格式");
            return;
        }

        const headerRow = rawScheduleData[headerRowIdx];
        const colToDayIdx = {}; // col -> 1(Mon)..7(Sun)
        headerRow.forEach((cell, idx) => {
            if (!cell) return;
            if (cell.includes('星期一')) colToDayIdx[idx] = 1;
            if (cell.includes('星期二')) colToDayIdx[idx] = 2;
            if (cell.includes('星期三')) colToDayIdx[idx] = 3;
            if (cell.includes('星期四')) colToDayIdx[idx] = 4;
            if (cell.includes('星期五')) colToDayIdx[idx] = 5;
            if (cell.includes('星期六')) colToDayIdx[idx] = 6;
            if (cell.includes('星期日') || cell.includes('星期天')) colToDayIdx[idx] = 7;
        });

        console.log("Day Map:", colToDayIdx);
        // Iterate rows below header
        const events = [];

        // Read current time settings from UI
        // Fix: Removed unused timeInputs variable
        const currentSlots = [];
        for (let i = 0; i < defaultTimeSlots.length; i++) {
            const startInput = document.querySelector(`input[data-idx="${i}"][data-type="start"]`);
            const endInput = document.querySelector(`input[data-idx="${i}"][data-type="end"]`);
            if (startInput && endInput) {
                currentSlots.push({ start: startInput.value, end: endInput.value });
            } else {
                currentSlots.push(defaultTimeSlots[i]);
            }
        }

        for (let r = headerRowIdx + 1; r < rawScheduleData.length; r++) {
            const row = rawScheduleData[r];
            if (!row || row.length === 0) continue;

            let periodNum = -1;

            // Helper to parse "第一节", "二", "3", etc.
            const parsePeriodCell = (cell) => {
                if (!cell) return -1;
                const s = String(cell).trim();

                // 1. Check for standard digits
                const digitMatch = s.match(/^(\d+)/);
                if (digitMatch) return parseInt(digitMatch[1]);

                // 2. Check for Chinese numerals
                const cnNums = {
                    '一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
                    '六': 6, '七': 7, '八': 8, '九': 9, '十': 10,
                    '十一': 11, '十二': 12
                };

                // Look for any key in string
                for (const [k, v] of Object.entries(cnNums)) {
                    if (s.includes(k)) return v;
                }

                return -1;
            };

            // Check first few columns (Index 0 and 1) for period info
            const p1 = parsePeriodCell(row[0]);
            const p2 = parsePeriodCell(row[1]);

            if (p2 !== -1) periodNum = p2;
            else if (p1 !== -1) periodNum = p1;

            if (periodNum === -1) periodNum = (r - headerRowIdx);

            Object.keys(colToDayIdx).forEach(colIdx => {
                const cellContent = row[colIdx];
                if (!cellContent) return;

                const courses = parseCourseString(cellContent);
                const dayOfWeeK = colToDayIdx[colIdx];

                courses.forEach(course => {
                    course.weeks.forEach(weekNum => {
                        const daysToAdd = (weekNum - 1) * 7 + (dayOfWeeK - 1);
                        const targetDate = new Date(semesterStart);
                        targetDate.setDate(semesterStart.getDate() + daysToAdd);

                        const pIdx = periodNum - 1;
                        const timeSlot = currentSlots[pIdx] || { start: '00:00', end: '00:00' };

                        // Double Period Logic (User Request: 1->1-2, 3->3-4 etc.)
                        let displayPeriod = `${periodNum}`;
                        let realEndTime = timeSlot.end;

                        // If odd period, assume it covers the next one too
                        if (periodNum % 2 !== 0) {
                            const nextPIdx = pIdx + 1;
                            if (currentSlots[nextPIdx]) {
                                realEndTime = currentSlots[nextPIdx].end;
                                displayPeriod = `${periodNum}-${periodNum + 1}`;
                            }
                        }

                        let timeOfDay = 'morning';
                        if (periodNum >= 5 && periodNum <= 8) timeOfDay = 'afternoon';
                        if (periodNum >= 9) timeOfDay = 'evening';

                        events.push({
                            date: targetDate,
                            period: periodNum,
                            displayPeriod: displayPeriod, // New Badge Text
                            startTime: timeSlot.start,
                            endTime: realEndTime, // Updated End Time
                            title: `${course.displayName} (${course.className})`,
                            location: course.location,
                            description: `${course.rawName} (${course.className})`,
                            timeOfDay: timeOfDay,
                            raw: course
                        });
                    });
                });
            });
        }

        generatedEvents = events;
        currentCalendarDate = new Date();
        renderMonthCalendar(currentCalendarDate);

        console.log("Schedule Generated!");
    } catch (error) {
        console.error(error);
        alert("Error generating schedule: " + error.message);
    }
}

// Refactored: Create DOM element for a month without attaching to DOM
function createMonthCalendarElement(date) {
    const year = date.getFullYear();
    const month = date.getMonth(); // 0-11

    const container = document.createElement('div');
    container.className = 'month-container'; // Wrapper for specific month

    // Header
    const header = document.createElement('h2');
    header.style.textAlign = 'center';
    header.textContent = `${year}年 ${month + 1}月`;
    header.className = 'month-title';
    container.appendChild(header);

    // Grid Header
    const grid = document.createElement('div');
    grid.className = 'calendar-grid';

    const weekDays = ['周一', '周二', '周三', '周四', '周五', '周六', '周日'];
    weekDays.forEach(d => {
        const h = document.createElement('div');
        h.className = 'calendar-header-cell';
        h.textContent = d;
        grid.appendChild(h);
    });

    // Days calculation
    const firstDayOfMonth = new Date(year, month, 1);
    let startDay = firstDayOfMonth.getDay();
    if (startDay === 0) startDay = 7;

    for (let i = 1; i < startDay; i++) {
        const empty = document.createElement('div');
        empty.className = 'calendar-day empty';
        grid.appendChild(empty);
    }

    const daysInMonth = new Date(year, month + 1, 0).getDate();

    for (let d = 1; d <= daysInMonth; d++) {
        const cell = document.createElement('div');
        cell.className = 'calendar-day';

        const currentDayDate = new Date(year, month, d);
        // Calculate Week Number
        // Week 1 starts on semesterStart (Monday). 
        // Calculate days difference from semesterStart
        const startInput = document.getElementById('semesterStart').value;
        let weekBadgeHtml = '';
        if (startInput) {
            // Fix: Manual parse to ensure local time 
            const parts = startInput.split('-');
            const startDate = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
            startDate.setHours(0, 0, 0, 0);

            // Current Date Normalized
            const cDate = new Date(year, month, d);
            cDate.setHours(0, 0, 0, 0);

            // Time diff in ms
            const diffTime = cDate.getTime() - startDate.getTime();
            const diffDays = Math.floor(diffTime / 86400000); // 1000*60*60*24

            // Logic: floor(diffDays / 7) + 1
            const weekNum = Math.floor(diffDays / 7) + 1;

            if (currentDayDate.getDay() === 1 || d === 1) {
                if (weekNum > 0 && weekNum <= 30) {
                    const hasEvents = generatedEvents.some(ev => ev.raw.weeks.includes(weekNum));
                    if (hasEvents) {
                        weekBadgeHtml = `<span class="week-badge">第${weekNum}周</span>`;
                    }
                }
            }
        }

        const dateNum = document.createElement('div');
        dateNum.className = 'day-number';
        // Flexbox: Badge | Number
        dateNum.innerHTML = `${weekBadgeHtml}<span>${d}</span>`;
        cell.appendChild(dateNum);

        // Find events
        const daysEvents = generatedEvents.filter(e =>
            e.date.getDate() === d &&
            e.date.getMonth() === month &&
            e.date.getFullYear() === year
        ).sort((a, b) => a.period - b.period);

        daysEvents.forEach(ev => {
            const evDiv = document.createElement('div');
            evDiv.className = `event-item type-${ev.timeOfDay}`;
            evDiv.title = `[第${ev.displayPeriod}节 ${ev.startTime}-${ev.endTime}]\n课程: ${ev.title}\n地点: ${ev.location}\n周次: ${ev.raw.weeks.join(',')}周`;

            evDiv.innerHTML = `
                <div>
                    <span class="ev-time">第${ev.displayPeriod}节</span>
                    <span class="ev-location">@${ev.location}</span>
                </div>
                <div class="ev-title">${ev.title}</div>
            `;
            cell.appendChild(evDiv);
        });

        grid.appendChild(cell);
    }

    container.appendChild(grid);
    return container;
}

function renderMonthCalendar(date) {
    const area = document.getElementById('calendarArea');
    area.innerHTML = '';

    const year = date.getFullYear();
    const month = date.getMonth();

    // Controls (Only for interactive view)
    const controls = document.createElement('div');
    controls.className = 'calendar-controls no-print';
    controls.innerHTML = `
        <button onclick="changeMonth(-1)">Previous</button>
        <span style="font-weight:bold; font-size:1.2rem;">${year}年 ${month + 1}月</span>
        <button onclick="changeMonth(1)">Next</button>
    `;
    area.appendChild(controls);

    // Use shared creator
    // But we need to hide the title inside the creator for interactive mode? 
    // Or just let it be. The user sees "202x年 X月" in controls.
    // Let's create the element.
    const calendarEl = createMonthCalendarElement(date);
    // Remove the title from the element if we already have controls? 
    // Or keep it. A title inside the grid area is fine.
    // For interactive view, remove the title from the created element as controls already have it.
    const titleInElement = calendarEl.querySelector('.month-title');
    if (titleInElement) {
        titleInElement.remove();
    }
    area.appendChild(calendarEl);
}

// Helper for color
function stringToColor(str) {
    let hash = 0;
    for (let i = 0; i < str.length; i++) {
        hash = str.charCodeAt(i) + ((hash << 5) - hash);
    }
    const c = (hash & 0x00FFFFFF).toString(16).toUpperCase();
    return '#' + '00000'.substring(0, 6 - c.length) + c;
}

window.changeMonth = function (delta) {
    currentCalendarDate.setMonth(currentCalendarDate.getMonth() + delta);
    renderMonthCalendar(currentCalendarDate);
}

// Print Handler
document.getElementById('btnPrint').addEventListener('click', () => {
    if (generatedEvents.length === 0) {
        alert("无日程数据");
        return;
    }

    // Calculate Date Range
    let minTime = Infinity;
    let maxTime = -Infinity;

    generatedEvents.forEach(e => {
        const t = e.date.getTime();
        if (t < minTime) minTime = t;
        if (t > maxTime) maxTime = t;
    });

    const startDate = new Date(minTime);
    const endDate = new Date(maxTime);

    // Align to Month start
    const startYear = startDate.getFullYear();
    const startMonth = startDate.getMonth();

    const endYear = endDate.getFullYear();
    const endMonth = endDate.getMonth();

    // Prepare Print View
    const area = document.getElementById('calendarArea');
    const OriginalHTML = area.innerHTML; // Backup
    area.innerHTML = ''; // Clear for print setup

    const printContainer = document.createElement('div');
    printContainer.className = 'print-all-container';

    let iterDate = new Date(startYear, startMonth, 1);
    // Loop until iterDate is past the end month
    while (iterDate.getFullYear() < endYear || (iterDate.getFullYear() === endYear && iterDate.getMonth() <= endMonth)) {
        const monthEl = createMonthCalendarElement(new Date(iterDate)); // Create a new date object to avoid mutation issues
        printContainer.appendChild(monthEl);

        // Next month
        iterDate.setMonth(iterDate.getMonth() + 1);
    }

    area.appendChild(printContainer);

    // Trigger Print
    window.print();

    // Restore (Optional, or leave it so they see what they printed)
    // A timeout helps to restore after print dialog closes (in some browsers)
    // But sticking to "Print View" is often less confusing.
    // Let's just reload the Current Month view to be safe/clean.
    setTimeout(() => {
        renderMonthCalendar(currentCalendarDate);
    }, 1000);
});

// ICS Export
document.getElementById('btnExport').addEventListener('click', () => {
    if (generatedEvents.length === 0) {
        alert("无日程数据");
        return;
    }

    const device = document.getElementById('exportTarget').value;

    // Header nuances
    // Windows/Outlook often prefers specific PRODID or METHOD
    let prodId = "-//ScheduleAG//CN";
    if (device === 'windows') prodId = "-//Microsoft Corporation//Outlook 16.0 MIMEDIR//EN";
    if (device === 'ios') prodId = "-//Apple Inc.//iOS 15.0//EN";

    let icsContent = `BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:${prodId}\r\nCALSCALE:GREGORIAN\r\nMETHOD:PUBLISH\r\n`;

    // Windows Outlook: Add TimeZone Definition? 
    // Simplify for now, usually VEVENT stats are enough.

    generatedEvents.forEach(ev => {
        // Format Date: YYYYMMDDTHHMMSS
        const dayStr = ev.date.toISOString().split('T')[0].replace(/-/g, '');
        const startStr = `${dayStr}T${ev.startTime.replace(/:/g, '')}00`;
        const endStr = `${dayStr}T${ev.endTime.replace(/:/g, '')}00`;

        let description = ev.description;
        if (device === 'ios') {
            // iOS sometimes likes cleaner description
        }

        icsContent += "BEGIN:VEVENT\r\n";
        icsContent += `UID:${Date.now()}-${Math.random()}@scheduleag\r\n`;
        icsContent += `DTSTAMP:${new Date().toISOString().replace(/[-:]/g, '').split('.')[0]}Z\r\n`;
        icsContent += `DTSTART;TZID=Asia/Shanghai:${startStr}\r\n`;
        icsContent += `DTEND;TZID=Asia/Shanghai:${endStr}\r\n`;
        icsContent += `SUMMARY:${ev.title}\r\n`;
        icsContent += `LOCATION:${ev.location}\r\n`;
        icsContent += `DESCRIPTION:${description}\r\n`;

        // Alarms
        if (device === 'ios' || device === 'android') {
            // 15 min reminder
            icsContent += "BEGIN:VALARM\r\nTRIGGER:-PT15M\r\nACTION:DISPLAY\r\nDESCRIPTION:Reminder\r\nEND:VALARM\r\n";
        }

        // Windows Outlook specific categories?
        if (device === 'windows') {
            const cat = ev.timeOfDay === 'morning' ? 'Blue Category' : (ev.timeOfDay === 'afternoon' ? 'Orange Category' : 'Purple Category');
            // icsContent += `CATEGORIES:${cat}\r\n`; // Outlook might need Master List, but safe to add
            icsContent += `X-MICROSOFT-CDO-BUSYSTATUS:BUSY\r\n`;
        }

        icsContent += "END:VEVENT\r\n";
    });

    icsContent += "END:VCALENDAR";

    const blob = new Blob([icsContent], { type: 'text/calendar;charset=utf-8' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `schedule_${device}.ics`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
});

// HTML Export
document.getElementById('btnSaveHtml').addEventListener('click', () => {
    if (generatedEvents.length === 0) {
        alert("无日程数据");
        return;
    }

    // 1. Calculate Date Range (Copy from Print logic)
    let minTime = Infinity;
    let maxTime = -Infinity;
    generatedEvents.forEach(e => {
        const t = e.date.getTime();
        if (t < minTime) minTime = t;
        if (t > maxTime) maxTime = t;
    });

    const startDate = new Date(minTime);
    const startYear = startDate.getFullYear();
    const startMonth = startDate.getMonth();

    const endDate = new Date(maxTime);
    const endYear = endDate.getFullYear();
    const endMonth = endDate.getMonth();

    // 2. Generate Content
    const container = document.createElement('div');
    container.className = 'print-all-container'; // Reuse print container class for layout

    let iterDate = new Date(startYear, startMonth, 1);
    while (iterDate.getFullYear() < endYear || (iterDate.getFullYear() === endYear && iterDate.getMonth() <= endMonth)) {
        const monthEl = createMonthCalendarElement(new Date(iterDate));
        container.appendChild(monthEl);
        iterDate.setMonth(iterDate.getMonth() + 1);
    }

    // 3. Get CSS logic
    // We need to inject the CSS. Since we can't easily read the file from JS in browser context 
    // without fetching, we will attempt to fetch 'style.css' or copy styles from document.styleSheets.
    // Simpler: Just fetch style.css assuming it's relative.
    // Or iterate stylesheets.
    let cssText = "";
    // Note: styles might be cross-origin protected if not local. But this is local file context or same origin.
    // Let's try to get all rules.
    for (let i = 0; i < document.styleSheets.length; i++) {
        try {
            const rules = document.styleSheets[i].cssRules;
            for (let j = 0; j < rules.length; j++) {
                cssText += rules[j].cssText + "\n";
            }
        } catch (e) { console.warn("Cannot read rules", e); }
    }

    // Add specific overrides for the HTML export to look good (similar to Print but scrollable)
    cssText += `
        body { background: white; font-family: 'Segoe UI', sans-serif; }
        .month-container { margin-bottom: 50px; border-bottom: 2px dashed #eee; padding-bottom: 20px; }
        .calendar-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 4px; }
        .calendar-day { border: 1px solid #eee; min-height: 100px; padding: 4px; }
        .calendar-header-cell { text-align: center; background: #f8f9fa; padding: 5px; font-weight: bold; }
        .event-item { margin-bottom: 2px; padding: 2px; font-size: 0.85em; cursor: help; }
        .type-morning { border-left: 3px solid #10b981; background: #ecfdf5; }
        .type-afternoon { border-left: 3px solid #f59e0b; background: #fffbeb; }
        .type-evening { border-left: 3px solid #8b5cf6; background: #f5f3ff; }
        h2 { text-align: center; color: #333; }
        /* Hide non-print stuff if copied from main css */
        .no-print { display: none; }
    `;

    const fullHtml = `
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>Course Schedule</title>
    <style>
        ${cssText}
    </style>
</head>
<body>
    <div style="max-width: 1200px; margin: 0 auto; padding: 20px;">
        ${container.innerHTML}
    </div>
</body>
</html>
    `;

    // 4. Download
    const blob = new Blob([fullHtml], { type: 'text/html;charset=utf-8' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `schedule_export.html`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
});
