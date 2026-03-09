
document.addEventListener('DOMContentLoaded', () => {
    const excelFileInput = document.getElementById('excel-file');
    const resultsTableDiv = document.getElementById('results-table');
    const exportBtn = document.getElementById('export-btn');

    let analysisResults = [];
    let lastRequiredCredits = {};

    const fileInputLabel = document.createElement('label');
    fileInputLabel.setAttribute('for', 'excel-file');
    fileInputLabel.textContent = '엑셀 파일 선택';
    excelFileInput.parentNode.insertBefore(fileInputLabel, excelFileInput.nextSibling);

    excelFileInput.addEventListener('change', (event) => {
        const uploadedFile = event.target.files[0];
        if (!uploadedFile) return;
        fileInputLabel.textContent = uploadedFile.name;
        startAnalysis(uploadedFile);
    });

    function startAnalysis(file) {
        console.clear();
        console.log("====== 데이터 분석 시작 ======");

        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const json_data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            console.log("1. 엑셀에서 읽어온 원본 데이터:", json_data);

            try {
                const requiredCredits = getRequiredCredits();
                lastRequiredCredits = requiredCredits;
                console.log("2. 사용자가 입력한 최소 이수 학점:", requiredCredits);
                const studentData = parseStudentData(json_data);
                console.log("4. 파싱 및 가공된 학생 데이터:", studentData);
                const results = analyzeStudentCredits(studentData, requiredCredits);
                console.log("5. 최종 분석 결과:", results);
                analysisResults = results;
                displayResults(results, requiredCredits);
                console.log("====== 데이터 분석 완료 ======");
            } catch (error) {
                console.error("분석 중 오류 발생:", error);
                alert(`오류가 발생했습니다: ${error.message}`);
                exportBtn.disabled = true;
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function getRequiredCredits() {
        return {
            '국어': parseInt(document.getElementById('korean-credit').value, 10),
            '수학': parseInt(document.getElementById('math-credit').value, 10),
            '영어': parseInt(document.getElementById('english-credit').value, 10),
            '한국사': parseInt(document.getElementById('history-credit').value, 10),
            '사회': parseInt(document.getElementById('social-credit').value, 10),
            '과학': parseInt(document.getElementById('science-credit').value, 10),
            '생활/교양': parseInt(document.getElementById('liberal-arts-credit').value, 10),
        };
    }

    function parseStudentData(data) {
        const students = {};
        let currentStudent = null;
        let currentYear = null;
        let currentSemester = null;
        let headerMapping = {};
        let headerRowFound = false;

        const headerAliases = {
            '번호': ['번 호'], '성명': ['성 명', '성  명'], '학년': ['학 년'], '학기': ['학 기'],
            '교과': ['교 과', ' 교 과 '], '과목': ['과 목'], '학점': ['학 점', '학점수']
        };

        function getCanonicalHeader(header) {
            if (header === null || typeof header === 'undefined') return null;
            const trimmedHeader = String(header).trim();
            for (const canonical in headerAliases) {
                if (canonical === trimmedHeader || headerAliases[canonical].includes(trimmedHeader)) {
                    return canonical;
                }
            }
            return trimmedHeader;
        }

        console.log("3. 학생 데이터 파싱 시작...");

        for (const row of data) {
            if (row.length === 0 || row.every(cell => cell === null || String(cell).trim() === '')) continue;

            const potentialHeaders = row.map(getCanonicalHeader);
            const isHeader = potentialHeaders.filter(h => h && ['번호', '성명', '교과', '학점'].includes(h)).length > 2;

            if (isHeader) {
                headerRowFound = true;
                headerMapping = {};
                row.forEach((header, index) => {
                    const canonicalHeader = getCanonicalHeader(header);
                    if (canonicalHeader) headerMapping[index] = canonicalHeader;
                });
                console.log("   - 헤더 매핑 정보:", headerMapping);
                continue;
            }

            if (!headerRowFound) continue;

            const rowData = {};
            row.forEach((cell, index) => {
                if (headerMapping[index]) rowData[headerMapping[index]] = cell;
            });

            if (String(row[0]).includes('이수학점 합계')) continue;

            if (rowData['성명'] && rowData['번호']) {
                currentStudent = String(rowData['성명']).trim();
                if (!students[currentStudent]) {
                    students[currentStudent] = { name: currentStudent, subjects: [] };
                }
            }

            if (rowData['학년']) currentYear = rowData['학년'];
            if (rowData['학기']) currentSemester = rowData['학기'];

            const creditsValue = rowData['학점'];

            if (currentStudent && rowData['교과'] && creditsValue) {
                let subjectCategory = String(rowData['교과']).trim();
                const subjectName = rowData['과목'] ? String(rowData['과목']).trim() : '';
                const credits = parseInt(creditsValue, 10);

                if (isNaN(credits)) continue;

                let classified = false;
                if (subjectName.includes('한국사')) {
                    subjectCategory = '한국사'; classified = true;
                } else if (['기술·가정/정보', '제2외국어', '교양', '기술·가정/정보/제2외국어/교양', '기술・가정/제2외국어/한문/교양'].some(c => subjectCategory.includes(c))) {
                    subjectCategory = '생활/교양'; classified = true;
                } else if (subjectCategory.includes('사회(역사/도덕포함)')) {
                    subjectCategory = '사회'; classified = true;
                } else if (['국어', '수학', '영어', '과학'].includes(subjectCategory)) {
                    classified = true;
                }

                students[currentStudent].subjects.push({
                    year: currentYear, semester: currentSemester, 
                    category: subjectCategory, name: subjectName, credits: credits, classified: classified
                });
            }
        }
        return Object.values(students);
    }

    function analyzeStudentCredits(studentData, requiredCredits) {
        return studentData.map(student => {
            const earnedCredits = { '국어': 0, '수학': 0, '영어': 0, '한국사': 0, '사회': 0, '과학': 0, '생활/교양': 0 };
            const unclassifiedSubjects = [];

            student.subjects.forEach(subject => {
                if (subject.classified && earnedCredits.hasOwnProperty(subject.category)) {
                    earnedCredits[subject.category] += subject.credits;
                } else if (!subject.classified) {
                    unclassifiedSubjects.push({ name: subject.name, credits: subject.credits });
                }
            });

            const results = {};
            for (const category in requiredCredits) {
                const required = requiredCredits[category];
                const earned = earnedCredits[category] || 0;
                results[category] = { earned, required, difference: earned - required };
            }
            return { name: student.name, results, unclassified: unclassifiedSubjects };
        });
    }

    function displayResults(results, requiredCredits) {
        resultsTableDiv.innerHTML = '';
        exportBtn.disabled = true;

        if (results.length === 0) {
            resultsTableDiv.innerHTML = '<p>분석할 데이터가 없거나, 데이터 형식이 올바르지 않습니다.</p>';
            return;
        }

        const container = document.createElement('div');
        container.className = 'results-table-container';
        const table = document.createElement('table');
        const thead = document.createElement('thead');
        const tbody = document.createElement('tbody');

        const headerRow = document.createElement('tr');
        const thName = document.createElement('th');
        thName.textContent = '학생명';
        headerRow.appendChild(thName);

        for (const category in requiredCredits) {
            const th = document.createElement('th');
            th.textContent = `${category} (${requiredCredits[category]})`;
            headerRow.appendChild(th);
        }
        const thUnclassified = document.createElement('th');
        thUnclassified.textContent = '미분류';
        headerRow.appendChild(thUnclassified);
        thead.appendChild(headerRow);
        table.appendChild(thead);

        results.forEach(studentResult => {
            const row = document.createElement('tr');
            const tdName = document.createElement('td');
            tdName.textContent = studentResult.name;
            row.appendChild(tdName);

            for (const category in studentResult.results) {
                const td = document.createElement('td');
                const result = studentResult.results[category];
                td.innerHTML = result.difference >= 0 ? `<span class="status-pass">충족 (${result.earned})</span>` : `<span class="status-fail">미충족 (${result.difference})</span>`;
                row.appendChild(td);
            }
            
            const tdUnclassified = document.createElement('td');
            tdUnclassified.textContent = studentResult.unclassified.map(s => `${s.name} (${s.credits})`).join(', ');
            row.appendChild(tdUnclassified);
            tbody.appendChild(row);
        });

        table.appendChild(tbody);
        container.appendChild(table);
        resultsTableDiv.appendChild(container);
        exportBtn.disabled = false;
    }

    function exportToExcel() {
        if (analysisResults.length === 0) {
            alert("내보낼 분석 결과가 없습니다.");
            return;
        }

        const categoryKeys = Object.keys(lastRequiredCredits);
        const headerRow = ['학생명', ...categoryKeys.map(cat => `${cat} (${lastRequiredCredits[cat]})`), '미분류'];

        const dataRows = analysisResults.map(studentResult => {
            const row = [studentResult.name];
            categoryKeys.forEach(category => {
                const result = studentResult.results[category];
                row.push(result.difference >= 0 ? `충족 (${result.earned})` : `미충족 (${result.difference})`);
            });
            row.push(studentResult.unclassified.map(s => `${s.name} (${s.credits})`).join(', '));
            return row;
        });

        const worksheet = XLSX.utils.aoa_to_sheet([headerRow, ...dataRows]);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, '분석결과');

        XLSX.writeFile(workbook, '이수_학점_분석_결과.xlsx');
    }

    exportBtn.addEventListener('click', exportToExcel);
});
