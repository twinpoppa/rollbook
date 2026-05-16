document.addEventListener('DOMContentLoaded', () => {
    // DOM Elements
    const fileDropArea = document.getElementById('file-drop-area');
    const fileInput = document.getElementById('excel-file');
    const resultSection = document.getElementById('result-section');
    const dashboardSection = document.getElementById('dashboard-section');
    const classListContainer = document.getElementById('class-list-container');
    const fileMessage = document.querySelector('.file-message');
    const startAttendanceBtn = document.getElementById('start-attendance-btn');
    
    const tabBtns = document.querySelectorAll('.tab-btn');
    const viewSections = document.querySelectorAll('.view-section');
    
    const classSelector = document.getElementById('class-selector');
    const studentGrid = document.getElementById('student-grid');
    const gradeSummaryGrid = document.getElementById('grade-summary-grid');
    const totalAbsentCount = document.getElementById('total-absent-count');

    const modalOverlay = document.getElementById('absent-modal');
    const modalTitle = document.getElementById('modal-title');
    const modalStudentList = document.getElementById('modal-student-list');
    const closeModalBtn = document.getElementById('close-modal-btn');

    // State
    let studentsByClass = {}; // Original data
    let attendanceState = {}; // { "1": { "1": "present", "2": "absent" } }

    // --- 1. File Upload & Parsing ---
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        fileDropArea.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) { e.preventDefault(); e.stopPropagation(); }

    ['dragenter', 'dragover'].forEach(eventName => {
        fileDropArea.addEventListener(eventName, () => fileDropArea.classList.add('dragover'), false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        fileDropArea.addEventListener(eventName, () => fileDropArea.classList.remove('dragover'), false);
    });

    fileDropArea.addEventListener('drop', (e) => {
        if (e.dataTransfer.files.length) {
            fileInput.files = e.dataTransfer.files;
            processFile(e.dataTransfer.files[0]);
        }
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length) processFile(e.target.files[0]);
    });

    function processFile(file) {
        fileMessage.textContent = '파일을 분석하는 중입니다...';
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                parseStudentData(jsonData);
            } catch (error) {
                console.error(error);
                alert('오류가 발생했습니다. 파일 형식을 확인해주세요.');
                fileMessage.textContent = '여기를 터치하거나 엑셀 파일을 끌어다 놓으세요';
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function parseStudentData(data) {
        studentsByClass = {};
        attendanceState = {};
        let studentCount = 0;

        data.forEach((row, index) => {
            if (!row || row.length < 3) return;

            const classNum = row[0];
            const studentNum = row[1];
            const name = row[2];
            
            if (index === 0 && isNaN(parseInt(classNum))) return;
            if (isNaN(parseInt(classNum))) return;

            if (!studentsByClass[classNum]) {
                studentsByClass[classNum] = [];
                attendanceState[classNum] = {};
            }

            const stNumInt = parseInt(studentNum);
            studentsByClass[classNum].push({
                studentNum: stNumInt,
                name: String(name).trim()
            });
            
            // 기본 출석 상태 설정
            attendanceState[classNum][stNumInt] = 'present';
            studentCount++;
        });

        // 학급별 학생 번호 순으로 정렬
        Object.keys(studentsByClass).forEach(cNum => {
            studentsByClass[cNum].sort((a, b) => a.studentNum - b.studentNum);
        });

        displayUploadResult(studentCount);
    }

    function displayUploadResult(totalCount) {
        document.querySelector('.upload-section').classList.add('hidden');
        resultSection.classList.remove('hidden');
        
        classListContainer.innerHTML = '';
        const sortedClasses = Object.keys(studentsByClass).sort((a, b) => parseInt(a) - parseInt(b));

        sortedClasses.forEach(classNum => {
            const div = document.createElement('div');
            div.className = 'class-item';
            div.innerHTML = `${classNum}반<span>총 ${studentsByClass[classNum].length}명</span>`;
            classListContainer.appendChild(div);
        });

        const sum = document.createElement('div');
        sum.style.gridColumn = '1 / -1';
        sum.style.textAlign = 'center';
        sum.style.color = 'var(--text-muted)';
        sum.innerHTML = `총 ${sortedClasses.length}개 반, ${totalCount}명의 학생 명단을 불러왔습니다.`;
        classListContainer.appendChild(sum);
        
        initDashboard(sortedClasses);
    }

    // --- 2. Dashboard Navigation ---
    startAttendanceBtn.addEventListener('click', () => {
        resultSection.classList.add('hidden');
        dashboardSection.classList.remove('hidden');
        renderGradeView(); // 초기 현황판 렌더링
    });

    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            tabBtns.forEach(b => b.classList.remove('active'));
            viewSections.forEach(v => v.classList.remove('active', 'hidden'));
            viewSections.forEach(v => v.classList.add('hidden'));
            
            btn.classList.add('active');
            const targetId = btn.getAttribute('data-target');
            document.getElementById(targetId).classList.remove('hidden');
            document.getElementById(targetId).classList.add('active');

            if(targetId === 'grade-view') {
                renderGradeView();
            }
        });
    });

    // --- 3. Dashboard Logic ---
    function initDashboard(sortedClasses) {
        // 반 선택 드롭다운 초기화
        classSelector.innerHTML = '<option value="">반을 선택하세요</option>';
        sortedClasses.forEach(cNum => {
            const opt = document.createElement('option');
            opt.value = cNum;
            opt.textContent = `${cNum}반`;
            classSelector.appendChild(opt);
        });

        classSelector.addEventListener('change', (e) => {
            renderClassView(e.target.value);
        });
    }

    // [우리 반 출결 체크 뷰]
    function renderClassView(classNum) {
        studentGrid.innerHTML = '';
        if (!classNum) return;

        const students = studentsByClass[classNum];
        
        students.forEach(st => {
            const status = attendanceState[classNum][st.studentNum];
            
            const card = document.createElement('div');
            card.className = `student-card ${status}`;
            card.innerHTML = `
                <span class="num">${st.studentNum}번</span>
                <span class="name">${st.name}</span>
            `;
            
            // 터치하여 출결 토글
            card.addEventListener('click', () => {
                const newStatus = status === 'present' ? 'absent' : 'present';
                attendanceState[classNum][st.studentNum] = newStatus;
                
                // 시각적 업데이트
                card.className = `student-card ${newStatus}`;
                
                // 전체 현황 데이터 갱신을 위해 백그라운드에서 renderGradeView 호출 (성능 최적화를 위해 생략하거나 필요시 호출)
            });

            studentGrid.appendChild(card);
        });
    }

    // [학년 전체 현황 뷰]
    function renderGradeView() {
        gradeSummaryGrid.innerHTML = '';
        let totalAbsent = 0;

        const sortedClasses = Object.keys(studentsByClass).sort((a, b) => parseInt(a) - parseInt(b));

        sortedClasses.forEach(classNum => {
            const classState = attendanceState[classNum];
            let absentCount = 0;
            let absentNames = [];

            studentsByClass[classNum].forEach(st => {
                if (classState[st.studentNum] === 'absent') {
                    absentCount++;
                    absentNames.push(st.name);
                }
            });

            totalAbsent += absentCount;

            const card = document.createElement('div');
            card.className = `grade-card ${absentCount > 0 ? 'has-absent' : ''}`;
            card.innerHTML = `
                <h3>${classNum}반</h3>
                <div class="absent-count">${absentCount}명</div>
                <div style="font-size:0.8rem; color:var(--text-muted); margin-top:5px;">결석(미도착)</div>
            `;

            card.addEventListener('click', () => {
                showAbsentModal(classNum, absentNames);
            });

            gradeSummaryGrid.appendChild(card);
        });

        totalAbsentCount.textContent = `총 ${totalAbsent}명 미도착`;
    }

    // 모달 관리
    function showAbsentModal(classNum, absentNames) {
        modalTitle.textContent = `${classNum}반 미도착자 명단 (${absentNames.length}명)`;
        modalStudentList.innerHTML = '';

        if (absentNames.length === 0) {
            modalStudentList.innerHTML = '<div style="text-align:center; padding: 20px; color:var(--text-muted);">미도착자가 없습니다! 👍</div>';
        } else {
            absentNames.forEach(name => {
                const item = document.createElement('div');
                item.className = 'modal-list-item';
                item.innerHTML = `<span style="font-weight:600;">${name}</span> <span style="color:var(--highlight); font-size:0.8rem;">미도착</span>`;
                modalStudentList.appendChild(item);
            });
        }

        modalOverlay.classList.remove('hidden');
    }

    closeModalBtn.addEventListener('click', () => {
        modalOverlay.classList.add('hidden');
    });

    modalOverlay.addEventListener('click', (e) => {
        if (e.target === modalOverlay) {
            modalOverlay.classList.add('hidden');
        }
    });
});
