document.addEventListener('DOMContentLoaded', () => {
    // === PÔVODNÉ ELEMENTY ===
    const fileInput = document.getElementById('file-input');
    const gridContainer = document.getElementById('grid-container');
    const btnRozdelovnik = document.getElementById('btn-rozdelovnik');

    // === NOVÉ ELEMENTY PRE DROP-ZONE ===
    const dropZone = document.getElementById('drop-zone');
    const fileNameDisplay = document.getElementById('file-name');
    
    // === PRIDANÝ ELEMENT TLAČIDLA "VYMAZAŤ" ===
    const btnClear = document.getElementById('btn-clear');

    // === PÔVODNÉ KONTROLY ===
    if (!fileInput) {
        console.error("Chyba: Element 'file-input' nebol nájdený.");
        return;
    }
    if (!gridContainer) {
        console.error("Chyba: Element 'grid-container' nebol nájdený.");
        return;
    }
    if (!btnRozdelovnik) {
        console.error("Chyba: Element 'btn-rozdelovnik' nebol nájdený.");
    }
    
    // === NOVÉ KONTROLY ===
    if (!dropZone) {
        console.error("Chyba: Element 'drop-zone' (label) nebol nájdený.");
        return;
    }
    if (!fileNameDisplay) {
        console.error("Chyba: Element 'file-name' nebol nájdený.");
        return;
    }
    // === KONTROLA PRE "VYMAZAŤ" ===
    if (!btnClear) {
        console.error("Chyba: Element 'btn-clear' nebol nájdený.");
    }

    let cisloSpisu = '';
    const originalGridContent = gridContainer.innerHTML; // Uložíme pôvodný obsah

    // === NOVÁ FUNKCIA PRE AKTUALIZÁCIU UI (NÁZVU SÚBORU) ===
    function updateFileName(file) {
        if (file) {
            // Zobrazí názov súboru
            fileNameDisplay.textContent = file.name;
            // Pridá triedu, ktorá skryje text "Presuňte súbor..."
            dropZone.classList.add('file-selected');
        } else {
            // Ak sa súbor zruší, vráti sa do pôvodného stavu
            fileNameDisplay.textContent = '';
            dropZone.classList.remove('file-selected');
        }
    }
    
    // === NOVÉ LISTENERY PRE DRAG & DROP ===

    // 1. Zabránenie predvolenému správaniu prehliadača
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    // 2. Pridanie vizuálneho zvýraznenia pri ťahaní súboru
    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => {
            dropZone.classList.add('drag-over');
        }, false);
    });

    // 3. Odstránenie vizuálneho zvýraznenia pri opustení zóny
    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('drag-over');
    }, false);

    // 4. Spracovanie pusteného súboru
    dropZone.addEventListener('drop', (e) => {
        dropZone.classList.remove('drag-over');
        
        const dt = e.dataTransfer;
        const files = dt.files;

        if (files.length > 0) {
            const file = files[0];
            // Overenie, či je to .xlsx (rovnako ako v inpute)
            if (file.name.endsWith('.xlsx')) {
                // Priradí súbor do nášho skrytého inputu (kvôli konzistencii)
                fileInput.files = files;
                // Aktualizuje zobrazený názov
                updateFileName(file);
                // Manuálne spustí vašu pôvodnú funkciu na spracovanie súboru
                // (musíme vytvoriť "falošný" event objekt)
                handleFile({ target: { files: files } });
            } else {
                alert('Prosím, nahrajte iba súbor vo formáte .xlsx');
                updateFileName(null); // Vyčistí názov súboru
            }
        }
    }, false);


    // === PÔVODNÁ FUNKCIA (handleFile) - BEZO ZMENY ===
    async function handleFile(event) {
        const file = event.target.files[0];
        if (!file) return;
        
        // JEDINÁ ZMENA: Ak je handleFile volaná z 'drop' eventu, 
        // fileInput.files už sú nastavené, ale pre 'change' event to
        // musíme spraviť tu. updateFileName() je teraz volané 
        // z 'change' listenera nižšie.

        gridContainer.innerHTML = '... Spracovávam súbor, prosím čakajte...';

        try {
            const workbook = await XlsxPopulate.fromDataAsync(await file.arrayBuffer());

            // NOVÉ: Zistiť aktívny hárok
            let sheet;
            try {
                const activeSheetIndex = workbook.activeSheet();
                if (typeof activeSheetIndex === 'number') {
                    sheet = workbook.sheet(activeSheetIndex);
                } else if (typeof activeSheetIndex === 'object' && activeSheetIndex !== null) {
                    sheet = activeSheetIndex;
                } else {
                    sheet = workbook.sheet(0);
                }
            } catch (e) {
                console.warn("Nepodarilo sa zistiť aktívny hárok, používam prvý hárok:", e.message);
                sheet = workbook.sheet(0);
            }

            console.log(`Načítaný hárok: ${sheet.name()}`);

            // NOVÉ: Zrušiť ukotvenie priečok (freeze panes)
            try {
                sheet.freezePanes(0, 0);
                console.log("Ukotvenie priečok zrušené.");
            } catch (e) {
                console.warn("Nepodarilo sa zrušiť ukotvenie priečok:", e.message);
            }

            // Načítanie čísla spisu
            try {
                const spisCell = sheet.cell("C3");
                cisloSpisu = (spisCell.value() === null || typeof spisCell.value() === 'undefined') ? '' : spisCell.value();
            } catch (e) {
                console.warn("Nepodarilo sa načítať číslo spisu z bunky C3:", e.message);
                cisloSpisu = '';
            }

            let dateHeaderText = '';
            try {
                const dateCell = sheet.cell("D1");
                dateHeaderText = (dateCell.value() === null || typeof dateCell.value() === 'undefined') ? '' : dateCell.value();
            } catch (e) {
                console.warn("Nepodarilo sa načítať hodnotu z bunky D1:", e.message);
            }

            let monthYearText = dateHeaderText;
            const match = dateHeaderText.match(/na mesiac\s+(.*)/i);
            if (match && match[1]) {
                monthYearText = match[1].trim();
            }

            const range = sheet.range("A13:AI64");
            const table = document.createElement('table');
            table.style.borderCollapse = 'collapse';
            table.style.width = '100%';
            const tbody = document.createElement('tbody');

            const numRows = 64 - 13 + 1;
            const numCols = 35;

            const headerRow = document.createElement('tr');
            const spisCell = document.createElement('td');
            spisCell.textContent = cisloSpisu || '';
            spisCell.setAttribute('colspan', '2');
            spisCell.style.fontWeight = 'bold';
            spisCell.style.padding = '8px';
            spisCell.style.backgroundColor = '#f0f0f0';
            spisCell.style.textAlign = 'left';
            headerRow.appendChild(spisCell);

            const dateCell = document.createElement('td');
            dateCell.textContent = monthYearText;
            dateCell.setAttribute('colspan', numCols - 3);
            dateCell.style.fontWeight = 'bold';
            dateCell.style.textAlign = 'center';
            dateCell.style.padding = '8px';
            dateCell.style.backgroundColor = '#f0f0f0';
            headerRow.appendChild(dateCell);

            tbody.appendChild(headerRow);

            function normalizeHexOrArgb(input) {
                if (!input) return null;
                if (typeof input === 'object') {
                    return null;
                }
                let hex = String(input).replace(/^#/, '').trim();
                if (hex.length === 8) {
                    const a = parseInt(hex.slice(0, 2), 16) / 255;
                    const r = parseInt(hex.slice(2, 4), 16);
                    const g = parseInt(hex.slice(4, 6), 16);
                    const b = parseInt(hex.slice(6, 8), 16);
                    if (a === 1) {
                        return `#${hex.slice(2)}`;
                    }
                    return `rgba(${r}, ${g}, ${b}, ${+a.toFixed(3)})`;
                }
                if (hex.length === 6) {
                    return `#${hex}`;
                }
                if (hex.length === 3) {
                    return `#${hex}`;
                }
                return `#${hex}`;
            }

            function extractColorFromFill(fill) {
                if (!fill) return null;
                let candidate = null;
                try {
                    if (fill.color) {
                        if (typeof fill.color.rgb === 'string') {
                            candidate = fill.color.rgb;
                        } else if (typeof fill.color.argb === 'string') {
                            candidate = fill.color.argb;
                        } else if (typeof fill.color === 'string') {
                            candidate = fill.color;
                        }
                    }
                    if (!candidate && fill.fgColor) {
                        if (typeof fill.fgColor.rgb === 'string') {
                            candidate = fill.fgColor.rgb;
                        } else if (typeof fill.fgColor.argb === 'string') {
                            candidate = fill.fgColor.argb;
                        } else if (typeof fill.fgColor === 'string') {
                            candidate = fill.fgColor;
                        }
                    }
                    if (!candidate && typeof fill.rgb === 'string') {
                        candidate = fill.rgb;
                    }
                    if (!candidate && typeof fill.argb === 'string') {
                        candidate = fill.argb;
                    }
                    if (!candidate && fill.color && fill.color.theme !== undefined) {
                        const theme = fill.color.theme;
                        const tint = parseFloat(fill.color.tint) || 0;
                        const themeColors = {
                            0: 'FFFFFF',
                            1: '000000',
                            2: 'E7E6E6',
                            3: '44546A',
                            4: '5B9BD5',
                            5: 'ED7D31',
                            6: 'A5A5A5',
                            7: 'FFC000'
                        };
                        let baseColor = themeColors[theme] || 'FFFFFF';
                        if (tint !== 0) {
                            const r = parseInt(baseColor.substring(0, 2), 16);
                            const g = parseInt(baseColor.substring(2, 4), 16);
                            const b = parseInt(baseColor.substring(4, 6), 16);
                            let newR, newG, newB;
                            if (tint < 0) {
                                newR = Math.round(r * (1 + tint));
                                newG = Math.round(g * (1 + tint));
                                newB = Math.round(b * (1 + tint));
                            } else {
                                newR = Math.round(r + (255 - r) * tint);
                                newG = Math.round(g + (255 - g) * tint);
                                newB = Math.round(b + (255 - b) * tint);
                            }
                            candidate = newR.toString(16).padStart(2, '0') +
                                newG.toString(16).padStart(2, '0') +
                                newB.toString(16).padStart(2, '0');
                            candidate = candidate.toUpperCase();
                        } else {
                            candidate = baseColor;
                        }
                    }
                    if (!candidate && typeof fill === 'string') {
                        candidate = fill;
                    }
                } catch (e) {
                    console.warn('Chyba pri extrakcii farby:', e);
                }
                return normalizeHexOrArgb(candidate);
            }

            function extractColorFromFont(font) {
                if (!font) return null;
                let candidate = null;
                try {
                    if (font.rgb) candidate = font.rgb;
                    if (!candidate && font.argb) candidate = font.argb;
                    if (!candidate && font.color) candidate = font.color.rgb || font.color.argb || font.color;
                    if (!candidate && typeof font === 'string') candidate = font;
                } catch (e) { }
                return normalizeHexOrArgb(candidate);
            }

            function safeCellStyle(cell, prop) {
                try {
                    return cell.style(prop);
                } catch (e) {
                    try {
                        const s = cell.style();
                        if (s && Object.prototype.hasOwnProperty.call(s, prop)) {
                            return s[prop];
                        }
                        return s;
                    } catch (e2) {
                        return null;
                    }
                }
            }

            for (let r = 0; r < numRows; r++) {
                const tr = document.createElement('tr');
                try {
                    const rowHeight = range.cell(r, 0).row().height();
                    if (rowHeight < 6) {
                        tr.classList.add("empty-excel-row");
                        tr.style.height = '3px';
                    }
                } catch (e) { }

                for (let c = 0; c < numCols; c++) {
                    if (c === 1) continue;

                    const cell = range.cell(r, c);
                    const td = document.createElement('td');
                    td.setAttribute('contenteditable', 'false');

                    const value = cell.value();
                    td.textContent = (value === null || typeof value === 'undefined') ? '' : value;

                    const fill = safeCellStyle(cell, "fill");
                    const bg = extractColorFromFill(fill);
                    if (bg) {
                        td.style.backgroundColor = bg;
                    }

                    const fontStyle = safeCellStyle(cell, "fontColor") || safeCellStyle(cell, "color") || safeCellStyle(cell, "font");
                    const fg = extractColorFromFont(fontStyle);
                    if (fg) {
                        td.style.color = fg;
                    }

                    if (c === 2) {
                        td.classList.add("name-column");
                    }

                    td.style.border = '1px solid #ccc';
                    td.style.padding = '4px';
                    td.style.whiteSpace = 'nowrap';
                    tr.appendChild(td);
                }
                tbody.appendChild(tr);
            }

            table.appendChild(tbody);
            gridContainer.innerHTML = '';
            gridContainer.appendChild(table);

        } catch (err) {
            console.error("Chyba pri spracovaní súboru XLSX:", err);
            gridContainer.innerHTML = `... Nastala chyba pri spracovaní súboru: ${err.message}`;
        }
    }

    // === PÔVODNÉ POMOCNÉ FUNKCIE (parseMonthYear, atď.) - BEZO ZMENY ===
    function parseMonthYear(text) {
        const monthsMap = {
            'január': 0, 'januára': 0, 'january': 0,
            'február': 1, 'februára': 1, 'february': 1,
            'marec': 2, 'marca': 2, 'march': 2,
            'apríl': 3, 'apríla': 3, 'april': 3,
            'máj': 4, 'mája': 4, 'may': 4,
            'jún': 5, 'júna': 5, 'june': 5,
            'júl': 6, 'júla': 6, 'july': 6,
            'august': 7, 'augusta': 7,
            'september': 8, 'septembra': 8,
            'október': 9, 'októbra': 9, 'october': 9,
            'november': 10, 'novembra': 10,
            'december': 11, 'decembra': 11
        };

        const normalized = text.toLowerCase().trim();
        for (const [monthName, monthIndex] of Object.entries(monthsMap)) {
            if (normalized.includes(monthName)) {
                const yearMatch = normalized.match(/\d{4}/);
                if (yearMatch) {
                    const year = parseInt(yearMatch[0]);
                    const monthNameCapitalized = monthName.charAt(0).toUpperCase() + monthName.slice(1);
                    return { month: monthNameCapitalized, year, monthIndex };
                }
            }
        }
        return null;
    }

    function getDaysInMonth(monthIndex, year) {
        return new Date(year, monthIndex + 1, 0).getDate();
    }

    function rgbToHex(rgb) {
        if (!rgb || !rgb.startsWith('rgb')) return '000000';
        const parts = rgb.match(/^rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)$/);
        if (!parts) return '000000';
        const r = parseInt(parts[1]).toString(16).padStart(2, '0');
        const g = parseInt(parts[2]).toString(16).padStart(2, '0');
        const b = parseInt(parts[3]).toString(16).padStart(2, '0');
        return (r + g + b).toUpperCase();
    }

    function extractSurname(fullName) {
        const words = fullName.split(/\s+/);
        const surnames = words.filter(word => {
            const clean = word.replace(/[.,;:]/g, '');
            return clean.length >= 2 && clean === clean.toUpperCase();
        });
        return surnames.length > 0 ? surnames.join(' ') : fullName;
    }

    function formatSurnameForNote(surname) {
        return surname.charAt(0).toUpperCase() + surname.slice(1).toLowerCase();
    }

    function isWeekend(day, monthIndex, year) {
        const date = new Date(year, monthIndex, day);
        const dayOfWeek = date.getDay();
        return dayOfWeek === 0 || dayOfWeek === 6;
    }

    // === UPRAVENÁ FUNKCIA (generateRozdelovnik) ===
    async function generateRozdelovnik() {
        console.log('Generujem rozdeľovník...');

        const gridContainer = document.getElementById('grid-container');
        const table = gridContainer.querySelector('table');

        if (!table) {
            alert('Najprv musíte nahrať a zobraziť plán služieb.');
            return;
        }

        try {
            const workbook = new ExcelJS.Workbook();
            const sheet = workbook.addWorksheet('Rozdeľovník');

            // Získanie čísla spisu a mesiaca z HTML tabuľky
            const headerCells = table.querySelectorAll('tr:first-child td');
            const cisloSpisuText = headerCells[0] ? headerCells[0].textContent.trim() : '';
            const titleCell = headerCells[1];

            if (!titleCell) {
                alert('Nepodarilo sa nájsť názov mesiaca a roka v tabuľke.');
                return;
            }

            const titleText = titleCell.textContent;
            const dateInfo = parseMonthYear(titleText);

            if (!dateInfo) {
                alert(`Nepodarilo sa z textu "${titleText}" extrahovať mesiac a rok.`);
                return;
            }

            const { month, year, monthIndex } = dateInfo;
            const numDays = getDaysInMonth(monthIndex, year);

            // Nastavenie šírok stĺpcov
            sheet.getColumn(1).width = 7;
            sheet.getColumn(2).width = 35;
            sheet.getColumn(3).width = 7;
            sheet.getColumn(4).width = 35;
            sheet.getColumn(5).width = 30;

            // Riadok 1: A1 = číslo spisu, E1 = "Dátum:"
            sheet.getCell('A1').value = cisloSpisuText;
            sheet.getCell('E1').value = 'Dátum:';

            // Riadok 3: Zlúčené A3:E3
            sheet.mergeCells('A3:E3');
            sheet.getCell('A3').value = `Rozdeľovník služieb operátorov na mesiac ${month} ${year}`;
            sheet.getCell('A3').alignment = { horizontal: 'center', vertical: 'middle' };
            sheet.getCell('A3').font = { bold: true, size: 14 };

            // Riadok 4: Zlúčené A4:E4
            sheet.mergeCells('A4:E4');
            sheet.getCell('A4').value = 'Koordinačného strediska IZS odboru krízového riadenia';
            sheet.getCell('A4').alignment = { horizontal: 'center', vertical: 'middle' };

            // Riadok 7: Hlavičky tabuľky
            sheet.getCell('A7').value = 'Dátum';
            sheet.getCell('B7').value = 'Denná zmena 06:30 - 18:30';
            sheet.getCell('C7').value = 'Dátum';
            sheet.getCell('D7').value = 'Nočná zmena 18:30 - 06:30';
            sheet.getCell('E7').value = 'Poznámka';

            // NASTAVENIE VÝŠKY RIADKA 7 NA 27
            sheet.getRow(7).height = 27;

            // DENNÁ ZMENA (B7) - Stredové zarovnanie + zelená výplň + veľkosť 14
            sheet.getCell('B7').alignment = { horizontal: 'center', vertical: 'top', wrapText: true };
            sheet.getCell('B7').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF92D050' } };
            sheet.getCell('B7').font = { bold: true, size: 14 };

            // NOČNÁ ZMENA (D7) - Stredové zarovnanie + oranžová výplň + veľkosť 14
            sheet.getCell('D7').alignment = { horizontal: 'center', vertical: 'top', wrapText: true };
            sheet.getCell('D7').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC000' } };
            sheet.getCell('D7').font = { bold: true, size: 14 };

            // Formátovanie ostatných hlavičiek (A7, C7, E7)
            ['A7', 'C7', 'E7'].forEach(cell => {
                sheet.getCell(cell).font = { bold: true };
                sheet.getCell(cell).alignment = { vertical: 'top', wrapText: true };
            });

            const firstDayColIndex = 2;
            const employeeRows = Array.from(table.querySelectorAll('tbody tr')).slice(1);

            if (employeeRows.length === 0) {
                alert('Nepodarilo sa nájsť riadky so zamestnancami v HTML tabuľke.');
                return;
            }

            console.log(`Počet riadkov zamestnancov: ${employeeRows.length}`);
            console.log(`Počet dní v mesiaci: ${numDays}`);

            const holidayDays = [];

            // Generovanie tabuľky od riadku 8
            for (let day = 1; day <= numDays; day++) {
                const excelRow = day + 7;
                const htmlColIndex = firstDayColIndex + (day - 1);

                let richTextSd = [];
                let richTextSn = [];
                let notesArray = [];
                let isHoliday = false;

                for (let rowIndex = 0; rowIndex < employeeRows.length; rowIndex++) {
                    const row = employeeRows[rowIndex];
                    if (!row.cells || row.cells.length < 2) continue;

                    const nameCell = row.cells[1];
                    if (!nameCell) continue;

                    const fullName = nameCell.textContent ? nameCell.textContent.trim() : '';
                    if (fullName === '') continue;

                    if (/^\d+$/.test(fullName) || fullName.toLowerCase().includes('meno')) {
                        continue;
                    }

                    const employeeName = extractSurname(fullName);
                    const shiftCell = row.cells[htmlColIndex];
                    if (!shiftCell) continue;

                    const cellBgColor = window.getComputedStyle(shiftCell).backgroundColor;
                    const yellowMatch = cellBgColor.match(/rgb\((\d+),\s*(\d+),\s*(\d+)\)/);
                    if (yellowMatch) {
                        const r = parseInt(yellowMatch[1]);
                        const g = parseInt(yellowMatch[2]);
                        const b = parseInt(yellowMatch[3]);
                        if (r > 240 && g > 240 && b < 50) {
                            isHoliday = true;
                        }
                    }

                    const shiftType = shiftCell.textContent ? shiftCell.textContent.trim().toLowerCase() : '';
                    if (shiftType === '') continue;

                    // === ZAČIATOK UPRAVENEJ LOGIKY ===
                    // Zmenená štruktúra: Najprv kontrolujeme farbu pozadia, 
                    // pretože má prednosť pred textom ('sd'/'sn').
                    
                    try {
                        const bgColor = window.getComputedStyle(shiftCell).backgroundColor;
                        const colorMatch = bgColor.match(/rgb\((\d+),\s*(\d+),\s*(\d+)\)/);
                        let hasBlueBackground = false;
                        let hasRedBackground = false;

                        if (colorMatch) {
                            const r = parseInt(colorMatch[1]);
                            const g = parseInt(colorMatch[2]);
                            const b = parseInt(colorMatch[3]);
                            hasBlueBackground = (r < 20 && g > 150 && g < 200 && b > 220);
                            hasRedBackground = (r > 200 && g < 100 && b < 100);
                        }

                        const formattedSurname = formatSurnameForNote(employeeName);

                        if (hasRedBackground) {
                            // PRAVIDLO 1: Červené pozadie -> IBA poznámka PN
                            // (Nepridá sa do zmeny, aj keby tam bolo 'sd'/'sn')
                            const noteText = `${formattedSurname}-PN`;
                            notesArray.push(noteText);

                        } else if (hasBlueBackground) {
                            // PRAVIDLO 2: Modré pozadie -> IBA poznámka D
                            // (Nepridá sa do zmeny, aj keby tam bolo 'sd'/'sn')
                            const noteText = `${formattedSurname}-D`;
                            notesArray.push(noteText);

                        } else if (shiftType === 'sd' || shiftType === 'sn') {
                            // PRAVIDLO 3: 'sd' alebo 'sn' (a nemá červené/modré pozadie)
                            // Pridá sa do zoznamu dennej alebo nočnej zmeny
                            
                            const shiftCellColor = window.getComputedStyle(shiftCell).color;
                            const hexColor = rgbToHex(shiftCellColor);
                            const nameFragment = {
                                text: employeeName,
                                font: { color: { argb: 'FF' + hexColor } }
                            };

                            if (shiftType === 'sd') {
                                if (richTextSd.length > 0) {
                                    richTextSd.push({ text: ', ', font: { color: { argb: 'FF000000' } } });
                                }
                                richTextSd.push(nameFragment);
                            } else if (shiftType === 'sn') {
                                if (richTextSn.length > 0) {
                                    richTextSn.push({ text: ', ', font: { color: { argb: 'FF000000' } } });
                                }
                                richTextSn.push(nameFragment);
                            }
                        
                        } else {
                            // PRAVIDLO 4: Iný text ('V', 'PN' bez farby, atď.)
                            // Pridá sa iba do poznámky
                            const noteText = `${formattedSurname}-${shiftType.toUpperCase()}`;
                            notesArray.push(noteText);
                        }

                    } catch (err) {
                        console.error(`Chyba pri spracovaní mena "${employeeName}":`, err);
                    }
                    // === KONIEC UPRAVENEJ LOGIKY ===
                }

                if (isHoliday) {
                    holidayDays.push(day);
                }

                sheet.getCell(`A${excelRow}`).value = day;
                sheet.getCell(`A${excelRow}`).alignment = { horizontal: 'center', vertical: 'top' };

                sheet.getCell(`C${excelRow}`).value = day;
                sheet.getCell(`C${excelRow}`).alignment = { horizontal: 'center', vertical: 'top' };

                if (isHoliday) {
                    ['A', 'B', 'C', 'D', 'E'].forEach(col => {
                        sheet.getCell(`${col}${excelRow}`).fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: 'FFFFFF00' }
                        };
                    });
                } else if (isWeekend(day, monthIndex, year)) {
                    sheet.getCell(`A${excelRow}`).fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFA5A5A5' }
                    };
                    sheet.getCell(`C${excelRow}`).fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFA5A5A5' }
                    };
                }

                if (richTextSd.length > 0) {
                    sheet.getCell(`B${excelRow}`).value = { richText: richTextSd };
                    sheet.getCell(`B${excelRow}`).alignment = { wrapText: true, vertical: 'top' };
                }

                if (richTextSn.length > 0) {
                    sheet.getCell(`D${excelRow}`).value = { richText: richTextSn };
                    sheet.getCell(`D${excelRow}`).alignment = { wrapText: true, vertical: 'top' };
                }

                if (notesArray.length > 0) {
                    sheet.getCell(`E${excelRow}`).value = notesArray.join(', ');
                    sheet.getCell(`E${excelRow}`).alignment = { wrapText: true, vertical: 'top' };
                }
            }

            // Riadok 39: Legendy
            sheet.getCell('A39').value = 'čierna farba - pravidelné striedanie služieb';
            sheet.getCell('A39').font = { bold: true };

            sheet.getCell('D39').value = 'modrá farba - služby za neprítomných';
            sheet.getCell('D39').font = { bold: true, color: { argb: 'FF00B0F0' } };

            // Riadok 42-44: Podpisy
            sheet.getCell('A42').value = 'Spracoval:';
            sheet.getCell('C42').value = 'Schvaľuje:';
            sheet.getCell('E42').value = 'Schvaľuje:';

            sheet.getCell('A43').value = 'Mgr. Juraj Tuhársky';
            sheet.getCell('C43').value = 'Mgr. Juraj Tuhársky';
            sheet.getCell('E43').value = 'Mgr. Mário Banič';

            sheet.getCell('A44').value = 'Ing. Silvia Sklenárová';
            sheet.getCell('C44').value = 'vedúci koordinačného strediska IZS';
            sheet.getCell('E44').value = 'vedúci odboru krízového riadenia';

            // Orámovania tabuľky (riadky 7 až numDays+7)
            const thinBorder = { style: 'thin', color: { argb: 'FF000000' } };
            const mediumBorder = { style: 'medium', color: { argb: 'FF000000' } };

            for (let row = 7; row <= numDays + 7; row++) {
                for (let col = 1; col <= 5; col++) {
                    const cell = sheet.getCell(row, col);
                    const border = {
                        top: thinBorder,
                        left: thinBorder,
                        bottom: thinBorder,
                        right: thinBorder
                    };

                    // Obvod tabuľky - medium border
                    if (row === 7) {
                        border.top = mediumBorder;
                    }
                    if (row === numDays + 7) {
                        border.bottom = mediumBorder;
                    }
                    if (col === 1) {
                        border.left = mediumBorder;
                    }
                    if (col === 5) {
                        border.right = mediumBorder;
                    }

                    // Hrubá čiara medzi stĺpcami
                    if (col === 1) {
                        border.right = mediumBorder;
                    }
                    if (col === 2) {
                        border.left = mediumBorder;
                        border.right = mediumBorder;
                    }
                    if (col === 3) {
                        border.left = mediumBorder;
                        border.right = mediumBorder;
                    }
                    if (col === 4) {
                        border.left = mediumBorder;
                        border.right = mediumBorder;
                    }
                    if (col === 5) {
                        border.left = mediumBorder;
                    }

                    cell.border = border;
                }
            }

            const fileName = `Rozdeľovník_${month}_${year}.xlsx`;
            const buffer = await workbook.xlsx.writeBuffer();
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

            if (window.navigator && window.navigator.msSaveOrOpenBlob) {
                window.navigator.msSaveOrOpenBlob(blob, fileName);
            } else {
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                document.body.appendChild(a);
                a.style = 'display: none';
                a.href = url;
                a.download = fileName;
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);
            }

            console.log('Rozdeľovník úspešne vygenerovaný!');

        } catch (err) {
            console.error('Chyba pri generovaní rozdeľovníka:', err);
            alert('Nastala chyba pri generovaní súboru: ' + err.message);
        }
    }

    // === UPRAVENÝ PÔVODNÝ LISTENER ===
    // Pôvodne: fileInput.addEventListener('change', handleFile);
    // Teraz:
    fileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) {
            updateFileName(file); // Najprv aktualizuj UI
            handleFile(e);        // Potom spusti spracovanie súboru
        } else {
            updateFileName(null); // Používateľ klikol na "Zrušiť"
        }
    });

    // === PÔVODNÝ LISTENER PRE TLAČIDLO (BEZO ZMENY) ===
    if (btnRozdelovnik) {
        btnRozdelovnik.addEventListener('click', generateRozdelovnik);
    }
    
    // === PRIDANÉ: LISTENER PRE TLAČIDLO "VYMAZAŤ" ===
    if (btnClear) {
        btnClear.addEventListener('click', () => {
            fileInput.value = null; // Vyčistí <input type="file">
            updateFileName(null); // Resetuje UI pre drop-zónu
            gridContainer.innerHTML = originalGridContent; // Vráti pôvodný text do gridu
            cisloSpisu = ''; // Resetuje globálnu premennú
            console.log('Aplikácia resetovaná.');
        });
    }
});