// 선생님이 전달해주신 Firebase 구성 정보
const firebaseConfig = {
  apiKey: "AIzaSyC025HS-8P6RBlFT_lSbc7A70lrf9CDTqA",
  authDomain: "rollbook-ffd93.firebaseapp.com",
  projectId: "rollbook-ffd93",
  storageBucket: "rollbook-ffd93.firebasestorage.app",
  messagingSenderId: "908996609938",
  appId: "1:908996609938:web:a7e70df0d11b7ef3e2e34b"
};

// Firebase 초기화 (Compat 버전 사용)
const app = firebase.initializeApp(firebaseConfig);
const db = firebase.firestore();

document.addEventListener('DOMContentLoaded', () => {
    // --- DOM Elements ---
    // Sections
    const loadingSection = document.getElementById('loading-section');
    const homeSection = document.getElementById('home-section');
    const newEventSection = document.getElementById('new-event-section');
    const setupSection = document.getElementById('setup-section');
    const resultSection = document.getElementById('result-section');
    const dashboardSection = document.getElementById('dashboard-section');
    
    // Buttons & Inputs (Home & Setup)
    const btnShowNewEvent = document.getElementById('btn-show-new-event');
    const btnShowJoinEvent = document.getElementById('btn-show-join-event');
    const btnCancelNewEvent = document.getElementById('btn-cancel-new-event');
    const btnCreateEvent = document.getElementById('btn-create-event');
    const newEventNameInput = document.getElementById('new-event-name');
    const btnFinishSetup = document.getElementById('btn-finish-setup');
    
    // File Uploads
    const setupFileDropArea = document.getElementById('setup-file-drop-area');
    const setupFileInput = document.getElementById('setup-excel-file');
    const updateFileDropArea = document.getElementById('update-file-drop-area');
    const updateFileInput = document.getElementById('update-excel-file');
    
    // Dashboard
    const dashboardEventTitle = document.getElementById('dashboard-event-title');
    const tabBtns = document.querySelectorAll('.tab-btn');
    const viewSections = document.querySelectorAll('.view-section');
    const classSelector = document.getElementById('class-selector');
    const studentGrid = document.getElementById('student-grid');
    const gradeSummaryGrid = document.getElementById('grade-summary-grid');
    const totalAbsentCount = document.getElementById('total-absent-count');
    const classListContainer = document.getElementById('class-list-container');

    // Modals
    const absentModal = document.getElementById('absent-modal');
    const modalTitle = document.getElementById('modal-title');
    const modalStudentList = document.getElementById('modal-student-list');
    const closeModalBtn = document.getElementById('close-modal-btn');
    
    const joinEventModal = document.getElementById('join-event-modal');
    const eventListContainer = document.getElementById('event-list-container');
    const closeJoinModalBtn = document.getElementById('close-join-modal-btn');
    
    const btnDashNewEvent = document.getElementById('btn-dash-new-event');
    const btnDashJoinEvent = document.getElementById('btn-dash-join-event');
    const btnDashManageEvent = document.getElementById('btn-dash-manage-event');
    
    const manageEventModal = document.getElementById('manage-event-modal');
    const manageEventListContainer = document.getElementById('manage-event-list-container');
    const closeManageModalBtn = document.getElementById('close-manage-modal-btn');

    // --- State ---
    let studentsByClass = {}; // Master List
    let attendanceState = {}; // Current Session Attendance
    let currentSessionId = null;
    let realtimeUnsubscribe = null;

    // --- Initialization ---
    async function initApp() {
        try {
            const masterRef = db.doc('settings/master_list');
            const snap = await masterRef.get();
            
            loadingSection.classList.add('hidden');
            
            if (snap.exists) {
                // 마스터 명단이 있는 경우
                studentsByClass = snap.data().studentsByClass || {};
                homeSection.classList.remove('hidden');
            } else {
                // 최초 사용 (명단 없음)
                setupSection.classList.remove('hidden');
            }
        } catch (error) {
            console.error("초기화 실패:", error);
            alert("데이터베이스 연결에 실패했습니다.");
        }
    }

    initApp();

    // --- File Upload & Parsing ---
    function setupDragAndDrop(dropArea, fileInput, callback) {
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            dropArea.addEventListener(eventName, preventDefaults, false);
        });

        function preventDefaults(e) { e.preventDefault(); e.stopPropagation(); }

        ['dragenter', 'dragover'].forEach(eventName => {
            dropArea.addEventListener(eventName, () => dropArea.classList.add('dragover'), false);
        });

        ['dragleave', 'drop'].forEach(eventName => {
            dropArea.addEventListener(eventName, () => dropArea.classList.remove('dragover'), false);
        });

        dropArea.addEventListener('drop', (e) => {
            if (e.dataTransfer.files.length) {
                fileInput.files = e.dataTransfer.files;
                callback(e.dataTransfer.files[0], dropArea);
            }
        });

        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length) callback(e.target.files[0], dropArea);
        });
    }

    function processExcelFile(file, dropArea) {
        const messageEl = dropArea.querySelector('.file-message');
        const originalMsg = messageEl.textContent;
        messageEl.textContent = '파일을 분석하는 중입니다...';
        
        const reader = new FileReader();
        reader.onload = async (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                let allData = [];
                workbook.SheetNames.forEach(sheetName => {
                    const worksheet = workbook.Sheets[sheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    
                    jsonData.forEach(row => {
                        allData.push({ row, sheetName });
                    });
                });
                
                await parseAndSaveMasterList(allData);
                
                messageEl.textContent = originalMsg; // 성공 후 원래 메시지로
            } catch (error) {
                console.error(error);
                alert('오류가 발생했습니다. 파일 형식을 확인해주세요.');
                messageEl.textContent = originalMsg;
            }
        };
        reader.readAsArrayBuffer(file);
    }

    async function parseAndSaveMasterList(data) {
        const newStudentsByClass = {};
        let studentCount = 0;

        data.forEach((item) => {
            const row = item.row;
            // 빈 줄 무시
            if (!row || row.length < 2) return;
            
            let classNum, studentNum, name;
            
            // 기존 폼 (A열: 반, B열: 번호, C열: 이름)
            if (row.length >= 3 && !isNaN(parseInt(row[0])) && !isNaN(parseInt(row[1]))) {
                classNum = row[0];
                studentNum = row[1];
                name = row[2];
            } 
            // 탭으로 구분된 폼 (A열: 번호, B열: 이름 - 반은 탭 이름에서 추출)
            else if (row.length >= 2 && !isNaN(parseInt(row[0]))) {
                studentNum = row[0];
                name = row[1];
                // 탭 이름에서 숫자 추출 (예: "1반" -> "1")
                const match = item.sheetName.match(/(\d+)/);
                if (match) {
                    classNum = match[1];
                } else {
                    return; // 반 번호를 찾을 수 없으면 무시
                }
            } else {
                return; // 헤더이거나 유효하지 않은 데이터 무시
            }

            if (isNaN(parseInt(classNum))) return;
            if (isNaN(parseInt(studentNum))) return;

            if (!newStudentsByClass[classNum]) {
                newStudentsByClass[classNum] = [];
            }

            const stNumInt = parseInt(studentNum);
            newStudentsByClass[classNum].push({
                studentNum: stNumInt,
                name: String(name).trim()
            });
            studentCount++;
        });

        // 정렬
        Object.keys(newStudentsByClass).forEach(cNum => {
            newStudentsByClass[cNum].sort((a, b) => a.studentNum - b.studentNum);
        });

        try {
            // 마스터 명단 업데이트
            const masterRef = db.doc('settings/master_list');
            await masterRef.set({
                studentsByClass: newStudentsByClass,
                updatedAt: new Date().toISOString()
            });
            
            studentsByClass = newStudentsByClass;
            
            // UI 업데이트
            if (!setupSection.classList.contains('hidden')) {
                // 초기 설정 화면인 경우
                setupSection.classList.add('hidden');
                resultSection.classList.remove('hidden');
                
                classListContainer.innerHTML = '';
                const sortedClasses = Object.keys(studentsByClass).sort((a, b) => parseInt(a) - parseInt(b));
                sortedClasses.forEach(classNum => {
                    const div = document.createElement('div');
                    div.className = 'class-item';
                    div.innerHTML = `${classNum}반<span>총 ${studentsByClass[classNum].length}명</span>`;
                    classListContainer.appendChild(div);
                });
            } else {
                // 대시보드에서 업데이트한 경우
                alert('명단이 성공적으로 업데이트되었습니다.');
                if (currentSessionId) {
                    renderGradeView(); // 바뀐 명단 반영 (출결 상태는 유지)
                }
            }
        } catch (error) {
            console.error("마스터 명단 저장 실패:", error);
            alert("명단 저장에 실패했습니다.");
        }
    }

    // 파일 업로드 이벤트 연결
    setupDragAndDrop(setupFileDropArea, setupFileInput, processExcelFile);
    setupDragAndDrop(updateFileDropArea, updateFileInput, processExcelFile);

    btnFinishSetup.addEventListener('click', () => {
        resultSection.classList.add('hidden');
        homeSection.classList.remove('hidden');
    });

    // --- Home Navigation ---
    btnShowNewEvent.addEventListener('click', () => {
        homeSection.classList.add('hidden');
        newEventSection.classList.remove('hidden');
    });

    btnCancelNewEvent.addEventListener('click', () => {
        newEventSection.classList.add('hidden');
        if (currentSessionId) {
            dashboardSection.classList.remove('hidden');
        } else {
            homeSection.classList.remove('hidden');
        }
    });
    
    if(btnDashNewEvent) {
        btnDashNewEvent.addEventListener('click', () => {
            dashboardSection.classList.add('hidden');
            newEventSection.classList.remove('hidden');
        });
    }
    
    if(btnDashJoinEvent) {
        btnDashJoinEvent.addEventListener('click', () => {
            btnShowJoinEvent.click();
        });
    }

    if(btnDashManageEvent) {
        btnDashManageEvent.addEventListener('click', async () => {
            alert("학년 부장 전용 메뉴입니다.");
            await loadManageEvents();
            manageEventModal.classList.remove('hidden');
        });
    }

    // --- Create New Event ---
    btnCreateEvent.addEventListener('click', async () => {
        const eventName = newEventNameInput.value.trim();
        if (!eventName) {
            alert('이벤트 이름을 입력해주세요.');
            return;
        }

        btnCreateEvent.textContent = '생성 중...';
        btnCreateEvent.disabled = true;

        try {
            // 마스터 명단을 기반으로 기본 출결 상태(present) 생성
            const initialAttendance = {};
            Object.keys(studentsByClass).forEach(cNum => {
                initialAttendance[cNum] = {};
                studentsByClass[cNum].forEach(st => {
                    initialAttendance[cNum][st.studentNum] = 'present';
                });
            });

            // 새 세션 문서 추가
            const sessionData = {
                eventName: eventName,
                createdAt: new Date().toISOString(),
                attendanceState: initialAttendance
            };

            const docRef = await db.collection('sessions').add(sessionData);
            
            currentSessionId = docRef.id;
            attendanceState = initialAttendance;
            dashboardEventTitle.textContent = `[${eventName}] 현황`;
            
            newEventSection.classList.add('hidden');
            dashboardSection.classList.remove('hidden');
            
            initDashboard();
        } catch (error) {
            console.error("이벤트 생성 실패:", error);
            alert("이벤트 생성에 실패했습니다.");
        } finally {
            btnCreateEvent.textContent = '생성 및 시작';
            btnCreateEvent.disabled = false;
        }
    });

    // --- Join Existing Event ---
    btnShowJoinEvent.addEventListener('click', async () => {
        const originalText = btnShowJoinEvent.innerHTML;
        btnShowJoinEvent.innerHTML = '불러오는 중...';
        btnShowJoinEvent.disabled = true;

        try {
            const querySnapshot = await db.collection('sessions').get();
            eventListContainer.innerHTML = '';
            
            if (querySnapshot.empty) {
                alert("현재 진행 중인 이벤트가 없습니다.");
                btnShowJoinEvent.innerHTML = originalText;
                btnShowJoinEvent.disabled = false;
                return;
            }

            // 날짜 최신순 정렬을 위해 배열로 변환
            const events = [];
            querySnapshot.forEach((doc) => {
                events.push({ id: doc.id, ...doc.data() });
            });
            events.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

            events.forEach(evt => {
                const item = document.createElement('div');
                item.className = 'event-list-item';
                
                const dateStr = new Date(evt.createdAt).toLocaleString('ko-KR', {
                    month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit'
                });

                item.innerHTML = `
                    <span class="event-name">${evt.eventName}</span>
                    <span class="event-date">${dateStr} 개설됨</span>
                `;
                
                item.addEventListener('click', () => joinSession(evt.id, evt.eventName, evt.attendanceState));
                eventListContainer.appendChild(item);
            });

            joinEventModal.classList.remove('hidden');
        } catch (error) {
            console.error("세션 목록 불러오기 실패:", error);
            alert("목록을 불러오지 못했습니다.");
        } finally {
            btnShowJoinEvent.innerHTML = originalText;
            btnShowJoinEvent.disabled = false;
        }
    });

    closeJoinModalBtn.addEventListener('click', () => {
        joinEventModal.classList.add('hidden');
    });

    function joinSession(sessionId, eventName, state) {
        currentSessionId = sessionId;
        attendanceState = state || {};
        dashboardEventTitle.textContent = `[${eventName}] 현황`;
        
        joinEventModal.classList.add('hidden');
        homeSection.classList.add('hidden');
        dashboardSection.classList.remove('hidden');
        
        initDashboard();
    }

    // --- Manage Events ---
    async function loadManageEvents() {
        try {
            const querySnapshot = await db.collection('sessions').get();
            manageEventListContainer.innerHTML = '';
            
            if (querySnapshot.empty) {
                manageEventListContainer.innerHTML = '<div style="text-align:center; padding: 20px; color:var(--text-muted);">진행 중인 이벤트가 없습니다.</div>';
                return;
            }

            const events = [];
            querySnapshot.forEach((doc) => {
                events.push({ id: doc.id, ...doc.data() });
            });
            events.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

            events.forEach(evt => {
                const item = document.createElement('div');
                item.className = 'event-list-item';
                item.style.display = 'flex';
                item.style.justifyContent = 'space-between';
                item.style.alignItems = 'center';
                
                const dateStr = new Date(evt.createdAt).toLocaleString('ko-KR', {
                    month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit'
                });

                item.innerHTML = `
                    <div>
                        <span class="event-name">${evt.eventName}</span>
                        <span class="event-date">${dateStr} 개설됨</span>
                    </div>
                    <button class="primary-button" style="background:var(--highlight); padding: 5px 10px; width: auto; font-size: 0.8rem; margin:0;">삭제</button>
                `;
                
                const deleteBtn = item.querySelector('button');
                deleteBtn.addEventListener('click', async (e) => {
                    e.stopPropagation();
                    if(confirm(`'${evt.eventName}' 이벤트를 정말 삭제하시겠습니까?`)) {
                        try {
                            await db.collection('sessions').doc(evt.id).delete();
                            alert("삭제되었습니다.");
                            item.remove();
                            if (currentSessionId === evt.id) {
                                dashboardSection.classList.add('hidden');
                                homeSection.classList.remove('hidden');
                                currentSessionId = null;
                            }
                        } catch(err) {
                            console.error("삭제 에러:", err);
                            alert("삭제에 실패했습니다.");
                        }
                    }
                });

                manageEventListContainer.appendChild(item);
            });
        } catch (error) {
            console.error("관리 목록 불러오기 실패:", error);
            alert("목록을 불러오지 못했습니다.");
        }
    }

    if(closeManageModalBtn) {
        closeManageModalBtn.addEventListener('click', () => {
            manageEventModal.classList.add('hidden');
        });
    }

    // --- Dashboard Tabs ---
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

    // --- Dashboard Logic ---
    function initDashboard() {
        const sortedClasses = Object.keys(studentsByClass).sort((a, b) => parseInt(a) - parseInt(b));
        
        classSelector.innerHTML = '<option value="">반을 선택하세요</option>';
        sortedClasses.forEach(cNum => {
            const opt = document.createElement('option');
            opt.value = cNum;
            opt.textContent = `${cNum}반`;
            classSelector.appendChild(opt);
        });

        // 탭이 바뀔때마다 이벤트 달아주지 않고 한 번만 달기 위해 기존 클론 또는 안전하게 달기
        classSelector.removeEventListener('change', onClassSelect);
        classSelector.addEventListener('change', onClassSelect);

        renderGradeView();
        setupRealtimeListener();
    }

    function onClassSelect(e) {
        renderClassView(e.target.value);
    }

    function setupRealtimeListener() {
        if (realtimeUnsubscribe) {
            realtimeUnsubscribe(); // 이전 리스너 해제
        }

        if (!currentSessionId) return;

        const sessionRef = db.collection('sessions').doc(currentSessionId);
        realtimeUnsubscribe = sessionRef.onSnapshot((docSnap) => {
            if (docSnap.exists) {
                const data = docSnap.data();
                if (data.attendanceState) {
                    attendanceState = data.attendanceState;
                    
                    const activeTab = document.querySelector('.tab-btn.active');
                    if (activeTab) {
                        const targetId = activeTab.getAttribute('data-target');
                        if (targetId === 'grade-view') {
                            renderGradeView();
                        } else if (targetId === 'class-view') {
                            renderClassView(classSelector.value);
                        }
                    }
                }
            }
        });
    }

    // [우리 반 출결 체크 뷰]
    function renderClassView(classNum) {
        studentGrid.innerHTML = '';
        if (!classNum) return;

        const students = studentsByClass[classNum] || [];
        
        students.forEach(st => {
            // 출결 상태가 없으면 기본값 present
            if(!attendanceState[classNum]) attendanceState[classNum] = {};
            const status = attendanceState[classNum][st.studentNum] || 'present';
            
            const card = document.createElement('div');
            card.className = `student-card ${status}`;
            card.innerHTML = `
                <span class="num">${st.studentNum}번</span>
                <span class="name">${st.name}</span>
            `;
            
            card.addEventListener('click', async () => {
                const currentStatus = attendanceState[classNum][st.studentNum] || 'present';
                const newStatus = currentStatus === 'present' ? 'absent' : 'present';
                
                attendanceState[classNum][st.studentNum] = newStatus;
                card.className = `student-card ${newStatus}`;
                
                try {
                    const sessionRef = db.collection('sessions').doc(currentSessionId);
                    await sessionRef.update({
                        [`attendanceState.${classNum}.${st.studentNum}`]: newStatus,
                        updatedAt: new Date().toISOString()
                    });
                } catch (error) {
                    console.error("출결 업데이트 실패:", error);
                    attendanceState[classNum][st.studentNum] = currentStatus;
                    card.className = `student-card ${currentStatus}`;
                    alert("서버 통신 중 오류가 발생했습니다.");
                }
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
            const classState = attendanceState[classNum] || {};
            let absentCount = 0;
            let absentStudents = [];

            (studentsByClass[classNum] || []).forEach(st => {
                const status = classState[st.studentNum] || 'present';
                if (status === 'absent') {
                    absentCount++;
                    absentStudents.push(`${st.studentNum}번 ${st.name}`);
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
                showAbsentModal(classNum, absentStudents);
            });

            gradeSummaryGrid.appendChild(card);
        });

        totalAbsentCount.textContent = `총 ${totalAbsent}명 미도착`;
    }

    // --- Modal Handling ---
    function showAbsentModal(classNum, absentStudents) {
        modalTitle.textContent = `${classNum}반 미도착자 명단 (${absentStudents.length}명)`;
        modalStudentList.innerHTML = '';

        if (absentStudents.length === 0) {
            modalStudentList.innerHTML = '<div style="text-align:center; padding: 20px; color:var(--text-muted);">미도착자가 없습니다! 👍</div>';
        } else {
            absentStudents.forEach(studentStr => {
                const item = document.createElement('div');
                item.className = 'modal-list-item';
                item.innerHTML = `<span style="font-weight:600;">${studentStr}</span> <span style="color:var(--highlight); font-weight:600;">미도착</span>`;
                modalStudentList.appendChild(item);
            });
        }
        absentModal.classList.remove('hidden');
    }

    closeModalBtn.addEventListener('click', () => {
        absentModal.classList.add('hidden');
    });

    [absentModal, joinEventModal, manageEventModal].forEach(modal => {
        if(modal) {
            modal.addEventListener('click', (e) => {
                if (e.target === modal) {
                    modal.classList.add('hidden');
                }
            });
        }
    });
});
