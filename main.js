
document.addEventListener('DOMContentLoaded', () => {
    const excelFileInput = document.getElementById('excel-file');
    const analyzeBtn = document.getElementById('analyze-btn');
    const resultsTableDiv = document.getElementById('results-table');

    let uploadedFile = null;

    // We need to add a label for the file input to make it clickable
    const fileInputLabel = document.createElement('label');
    fileInputLabel.setAttribute('for', 'excel-file');
    fileInputLabel.textContent = '엑셀 파일 선택';
    excelFileInput.parentNode.insertBefore(fileInputLabel, excelFileInput.nextSibling);


    excelFileInput.addEventListener('change', (event) => {
        uploadedFile = event.target.files[0];
        if (uploadedFile) {
            fileInputLabel.textContent = uploadedFile.name;
            analyzeBtn.disabled = false;
        } else {
            fileInputLabel.textContent = '엑셀 파일 선택';
            analyzeBtn.disabled = true;
        }
    });

    analyzeBtn.addEventListener('click', () => {
        if (!uploadedFile) {
            alert("엑셀 파일을 먼저 업로드해주세요.");
            return;
        }

        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const json_data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            try {
                const requiredCredits = getRequiredCredits();
                const studentData = parseStudentData(json_data);
                const results = analyzeStudentCredits(studentData, requiredCredits);
                displayResults(results, requiredCredits);
            } catch (error) {
                console.error("분석 중 오류 발생:", error);
                alert(`오류가 발생했습니다: ${error.message}`);
            }
        };
        reader.readAsArrayBuffer(uploadedFile);
    });

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
        let headerRowFound = false;
        let headerMapping = {};

        for (const row of data) {
            // 헤더 행 찾기 ('번 호', '성 명' 등이 포함된 행)
            if (row.includes('번 호') && row.includes('성 명')) {
                headerRowFound = true;
                // 헤더 매핑 생성 (A, B, C... -> 번 호, 성 명...)
                row.forEach((header, index) => {
                  if(header) {
                    headerMapping[index] = header.trim();
                  }
                });
                continue;
            }

            if (!headerRowFound) continue;

            // 데이터 행 처리
            const rowData = {};
            row.forEach((cell, index) => {
              const header = headerMapping[index];
              if (header) {
                rowData[header] = cell;
              }
            });


            // 학생 이름 처리
            if (rowData['성 명']) {
                currentStudent = rowData['성 명'].trim();
                if (!students[currentStudent]) {
                    students[currentStudent] = { name: currentStudent, subjects: [] };
                }
            }
            
            // 학년, 학기 처리
            if(rowData['학 년']) currentYear = rowData['학 년'];
            if(rowData['학 기']) currentSemester = rowData['학 기'];

            const creditsValue = rowData['학 점'] || rowData['학점수'];
            if (currentStudent && rowData['교 과'] && creditsValue) {
                let subjectCategory = rowData['교 과'].trim();
                const subjectName = rowData['과 목'] ? rowData['과 목'].trim() : '';
                const credits = parseInt(creditsValue, 10);

                if (isNaN(credits)) continue;

                // 교과 분류 규칙 적용
                if (subjectName.includes('한국사')) {
                    subjectCategory = '한국사';
                } else if (['기술·가정/정보', '제2외국어', '교양', '기술·가정/정보/제2외국어/교양'].some(c => subjectCategory.includes(c))) {
                    subjectCategory = '생활/교양';
                } else if (subjectCategory.includes('사회(역사/도덕포함)')) {
                    subjectCategory = '사회';
                }


                students[currentStudent].subjects.push({
                    year: currentYear,
                    semester: currentSemester,
                    category: subjectCategory,
                    name: subjectName,
                    credits: credits
                });
            }
        }
        return Object.values(students);
    }

    function analyzeStudentCredits(studentData, requiredCredits) {
        return studentData.map(student => {
            const earnedCredits = {
                '국어': 0, '수학': 0, '영어': 0, '한국사': 0,
                '사회': 0, '과학': 0, '생활/교양': 0
            };

            student.subjects.forEach(subject => {
                if (earnedCredits.hasOwnProperty(subject.category)) {
                    earnedCredits[subject.category] += subject.credits;
                }
            });

            const results = {};
            let allPassed = true;
            for (const category in requiredCredits) {
                const required = requiredCredits[category];
                const earned = earnedCredits[category] || 0;
                const difference = earned - required;
                results[category] = { earned, required, difference };
                if (difference < 0) allPassed = false;
            }
            return { name: student.name, results, allPassed };
        });
    }

    function displayResults(results, requiredCredits) {
        resultsTableDiv.innerHTML = '';
        if (results.length === 0) {
            resultsTableDiv.innerHTML = '<p>분석할 데이터가 없거나, 데이터 형식이 올바르지 않습니다.</p>';
            return;
        }

        const container = document.createElement('div');
        container.className = 'results-table-container';
        const table = document.createElement('table');
        const thead = document.createElement('thead');
        const tbody = document.createElement('tbody');

        // 테이블 헤더 생성
        const headerRow = document.createElement('tr');
        const thName = document.createElement('th');
        thName.textContent = '학생명';
        headerRow.appendChild(thName);

        for (const category in requiredCredits) {
            const th = document.createElement('th');
            th.textContent = `${category} (${requiredCredits[category]})`;
            headerRow.appendChild(th);
        }
        thead.appendChild(headerRow);
        table.appendChild(thead);

        // 테이블 바디 생성
        results.forEach(studentResult => {
            const row = document.createElement('tr');
            const tdName = document.createElement('td');
            tdName.textContent = studentResult.name;
            row.appendChild(tdName);

            for (const category in studentResult.results) {
                const td = document.createElement('td');
                const result = studentResult.results[category];
                
                if (result.difference >= 0) {
                    td.innerHTML = `<span class="status-pass">충족 (${result.earned})</span>`;
                } else {
                    td.innerHTML = `<span class="status-fail">미충족 (${result.difference})</span>`;
                }
                row.appendChild(td);
            }
            tbody.appendChild(row);
        });

        table.appendChild(tbody);
        container.appendChild(table);
        resultsTableDiv.appendChild(container);
    }
});
