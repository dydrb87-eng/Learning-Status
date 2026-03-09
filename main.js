document.addEventListener('DOMContentLoaded', () => {
    const excelFileInput = document.getElementById('excel-file');
    const resultsTableDiv = document.getElementById('results-table');
    const exportBtn = document.getElementById('export-btn');

    const unclassifiedContainer = document.getElementById('unclassified-subjects-container');
    const unclassifiedTitle = document.getElementById('unclassified-title');
    const unclassifiedList = document.getElementById('unclassified-list');

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
            
            console.log("1. 엑셀 원본 데이터:", json_data);

            try {
                const requiredCredits = getRequiredCredits();
                lastRequiredCredits = requiredCredits;
                console.log("2. 필수 이수 학점 기준:", requiredCredits);
                const studentData = parseStudentData(json_data);
                console.log("4. 최종 가공된 학생 데이터:", studentData);
                const results = analyzeStudentCredits(studentData, requiredCredits);
                console.log("5. 최종 분석 결과:", results);
                analysisResults = results;
                displayResults(results, requiredCredits);
                console.log("====== 데이터 분석 완료 ======");
            } catch (error) {
                console.error("분석 중 치명적 오류 발생:", error);
                alert(`오류가 발생했습니다: ${error.message}`);
                exportBtn.disabled = true;
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function getRequiredCredits() {
        const creditIds = {
            '국어': 'korean-credit', '수학': 'math-credit', '영어': 'english-credit',
            '한국사': 'history-credit', '사회': 'social-credit', '과학': 'science-credit',
            '체육': 'sports-credit', '예술': 'art-credit', '생활/교양': 'liberal-arts-credit',
        };

        const requiredCredits = {};
        for (const [category, id] of Object.entries(creditIds)) {
            const element = document.getElementById(id);
            if (element) {
                requiredCredits[category] = parseInt(element.value, 10);
            } else {
                console.warn(`[경고] HTML에서 학점 입력 필드(ID: #${id})를 찾을 수 없습니다. 기본값 10을 사용합니다.`);
                requiredCredits[category] = 10; 
            }
        }
        return requiredCredits;
    }

    function parseStudentData(data) {
        console.log("3. 학생 데이터 파싱 최종 로직 (매 행 헤더 검사) 시작...");
        const students = {};
        let headerMapping = {};
        let currentStudentName = "";
        let headerFoundAtLeastOnce = false;

        const headerAliases = {
            '성명': ['성 명', '성  명'],
            '교과': ['교 과', ' 교 과 '],
            '과목': ['과 목'],
            '학점': ['학 점', '학점수']
        };

        data.forEach((row, rowIndex) => {
            if (!row || row.length === 0 || row.every(cell => !cell || String(cell).trim() === '')) {
                return;
            }

            const potentialHeaders = {};
            row.forEach((cell, index) => {
                const trimmedCell = String(cell).trim();
                for (const canonical in headerAliases) {
                    if (trimmedCell === canonical || headerAliases[canonical].includes(trimmedCell)) {
                        potentialHeaders[canonical] = index;
                    }
                }
            });

            const isHeader = potentialHeaders.hasOwnProperty('성명') && potentialHeaders.hasOwnProperty('교과') && potentialHeaders.hasOwnProperty('학점');

            if (isHeader) {
                headerMapping = potentialHeaders;
                headerFoundAtLeastOnce = true;
                currentStudentName = "";
                console.log(`   - ${rowIndex + 1}번째 행 => 헤더 발견/갱신. 열 매핑:`, headerMapping);
                return;
            }

            if (!headerFoundAtLeastOnce) {
                return;
            }
            
            const nameCell = row[headerMapping['성명']];
            if (nameCell && String(nameCell).trim() !== '') {
                currentStudentName = String(nameCell).trim();
            }

            if (!currentStudentName) {
                return;
            }

            const originalCategory = row[headerMapping['교과']];
            const creditsValue = row[headerMapping['학점']];

            if (typeof originalCategory === 'undefined' || typeof creditsValue === 'undefined' || creditsValue === null) {
                return; 
            }

            const credits = parseInt(creditsValue, 10);
            if (isNaN(credits)) {
                return;
            }

            if (!students[currentStudentName]) {
                students[currentStudentName] = { name: currentStudentName, subjects: [] };
            }
            
            const subjectName = row[headerMapping['과목']] ? String(row[headerMapping['과목']]).trim() : '';
            let finalCategory = String(originalCategory).trim();

            if (subjectName.includes('한국사')) {
                finalCategory = '한국사';
            } else if (finalCategory.includes('사회')) {
                finalCategory = '사회';
            } else if (['기술·가정', '정보', '제2외국어', '한문', '교양'].some(c => finalCategory.includes(c))) {
                finalCategory = '생활/교양';
            }

            const definedCategories = ['국어', '수학', '영어', '한국사', '사회', '과학', '체육', '예술', '생활/교양'];
            const classified = definedCategories.includes(finalCategory);

            students[currentStudentName].subjects.push({
                originalCategory: String(originalCategory).trim(),
                finalCategory: finalCategory,
                name: subjectName,
                credits: credits,
                classified: classified
            });
        });

        return Object.values(students);
    }

    function analyzeStudentCredits(studentData, requiredCredits) {
        const categoryKeys = Object.keys(requiredCredits);

        return studentData.map(student => {
            const earnedCredits = {};
            categoryKeys.forEach(key => earnedCredits[key] = 0);
            const unclassifiedSubjects = [];
            
            student.subjects.forEach(subject => {
                if (subject.classified && earnedCredits.hasOwnProperty(subject.finalCategory)) {
                    earnedCredits[subject.finalCategory] += subject.credits;
                } else if (!subject.classified) {
                    unclassifiedSubjects.push({ 
                        name: subject.name, 
                        category: subject.originalCategory 
                    });
                }
            });

            const results = {};
            categoryKeys.forEach(category => {
                const required = requiredCredits[category];
                const earned = earnedCredits[category] || 0;
                results[category] = { earned, required, difference: earned - required };
            });

            const kmeTotal = (results['국어']?.earned || 0) + (results['수학']?.earned || 0) + (results['영어']?.earned || 0);

            return { 
                name: student.name, 
                results, 
                kmeTotal: kmeTotal, 
                unclassifiedSubjects 
            };
        });
    }

    function displayResults(results, requiredCredits) {
        resultsTableDiv.innerHTML = '';
        exportBtn.disabled = true;
        unclassifiedContainer.hidden = true;
        unclassifiedList.innerHTML = '';

        if (results.length === 0) {
            resultsTableDiv.innerHTML = '<p>분석할 데이터가 없거나, 데이터 형식이 올바르지 않습니다. 헤더(성명, 교과, 학점)를 확인해주세요.</p>';
            return;
        }

        const container = document.createElement('div');
        container.className = 'results-table-container';
        const table = document.createElement('table');
        const thead = document.createElement('thead');
        const tbody = document.createElement('tbody');

        const headerRow = document.createElement('tr');
        headerRow.innerHTML = '<th>학생명</th>' + 
            Object.keys(requiredCredits).map(cat => `<th>${cat} (${requiredCredits[cat]})</th>`).join('') + 
            '<th>국영수합</th>';
        thead.appendChild(headerRow);
        table.appendChild(thead);

        const allUnclassified = [];

        results.forEach(studentResult => {
            const row = document.createElement('tr');
            let rowHTML = `<td>${studentResult.name}</td>`;
            
            Object.keys(studentResult.results).forEach(category => {
                const result = studentResult.results[category];
                const statusClass = result.difference >= 0 ? 'status-pass' : 'status-fail';
                const text = result.difference >= 0 ? `충족 (${result.earned})` : `미충족 (${result.difference})`;
                rowHTML += `<td><span class="${statusClass}">${text}</span></td>`;
            });

            const kmePercentage = ((studentResult.kmeTotal / 174) * 100).toFixed(1);
            rowHTML += `<td>${studentResult.kmeTotal}/174 (${kmePercentage}%)</td>`;
            
            row.innerHTML = rowHTML;
            tbody.appendChild(row);

            if (studentResult.unclassifiedSubjects.length > 0) {
                const studentUnclassified = studentResult.unclassifiedSubjects.map(s => `${s.name}(${s.category})`).join(', ');
                allUnclassified.push({ name: studentResult.name, subjects: studentUnclassified });
            }
        });

        table.appendChild(tbody);
        container.appendChild(table);
        resultsTableDiv.appendChild(container);

        if (allUnclassified.length > 0) {
            unclassifiedContainer.hidden = false;
            const totalCount = allUnclassified.reduce((acc, curr) => acc + curr.subjects.split(', ').length, 0);
            unclassifiedTitle.textContent = `미분류 과목 (${totalCount}개)`;
            allUnclassified.forEach(item => {
                const li = document.createElement('li');
                li.innerHTML = `<strong>${item.name}:</strong> ${item.subjects}`;
                unclassifiedList.appendChild(li);
            });
        }

        exportBtn.disabled = false;
    }

    function exportToExcel() {
        if (analysisResults.length === 0) {
            alert("내보낼 분석 결과가 없습니다.");
            return;
        }

        const categoryKeys = Object.keys(lastRequiredCredits);
        const headerRow = ['학생명', ...categoryKeys.map(cat => `${cat} (${lastRequiredCredits[cat]})`), '국영수합'];

        const dataRows = analysisResults.map(studentResult => {
            const row = [studentResult.name];
            categoryKeys.forEach(category => {
                const result = studentResult.results[category];
                row.push(result.difference >= 0 ? `충족 (${result.earned})` : `미충족 (${result.difference})`);
            });
            const kmePercentage = ((studentResult.kmeTotal / 174) * 100).toFixed(1);
            row.push(`${studentResult.kmeTotal}/174 (${kmePercentage}%)`);
            return row;
        });

        const worksheet = XLSX.utils.aoa_to_sheet([headerRow, ...dataRows]);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, '분석결과');

        XLSX.writeFile(workbook, '이수_학점_분석_결과.xlsx');
    }

    exportBtn.addEventListener('click', exportToExcel);
});