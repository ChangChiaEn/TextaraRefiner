document.addEventListener('DOMContentLoaded', () => {
    // ⚠️ 請將此處的 URL 替換為您 Colab 後端生成的 ngrok 網址
    const BACKEND_URL = 'https://7febba17fd34.ngrok-free.app';

    // --- DOM Elements ---
    const reportListEl = document.getElementById('report-list');
    const reportListLoaderEl = document.getElementById('report-list-loader');
    const welcomeScreenEl = document.getElementById('welcome-screen');
    const reportViewEl = document.getElementById('report-view');
    const reportTitleEl = document.getElementById('report-title');
    const reportDisplayEl = document.getElementById('report-display');
    const saveReportBtn = document.getElementById('save-report-btn');
    const refinePanelEl = document.getElementById('refine-panel');
    const selectedTextPreviewEl = document.getElementById('selected-text-preview');
    const chatHistoryEl = document.getElementById('chat-history');
    const refineInstructionInput = document.getElementById('refine-instruction');
    const sendInstructionBtn = document.getElementById('send-instruction-btn');
    const connectionStatusEl = document.getElementById('connection-status');
    const saveStatusEl = document.getElementById('save-status-message');
    
    // --- State ---
    let ws = null;
    let currentReport = { id: null, name: null };
    let currentSelection = { range: null, text: '' };
    let chatHistory = [];
    let isApplyingChange = false;
    let isLoadingReport = false;
    let selectReportDebounced = null;
    let refinedSections = [];

    // --- Initialization & Connection ---
    function init() {
        fetchReports();
        setupWebSocket();
        addEventListeners();
        handleUrlParameters();
    }


    async function handleUrlParameters() {
        const params = new URLSearchParams(window.location.search);
        const reportName = params.get('report');

        if (!reportName) {
            return; 
        }

        console.log(`URL 參數請求加載報告: ${reportName}`);

        welcomeScreenEl.classList.add('hidden');
        reportViewEl.classList.remove('hidden');
        reportDisplayEl.innerHTML = `<div class="loader"></div><p style="text-align:center;">正在尋找並加載報告: ${reportName}...</p>`;
        
        const findAndLoadReport = async () => {
            const reportListItems = reportListEl.querySelectorAll('li[data-name]');
            if (reportListItems.length === 0 && !reportListEl.querySelector('.error')) {
                // 如果列表是空的且沒有錯誤，稍後再試
                setTimeout(findAndLoadReport, 500);
                return;
            }

            let found = false;
            for (const item of reportListItems) {
                const itemName = item.dataset.name.split('/').pop().trim();
                if (itemName === reportName) {
                    console.log(`找到匹配的報告: ${item.dataset.name}, ID: ${item.dataset.id}`);
                    found = true;
                    // 模擬點擊來選中並加載報告
                    item.click(); 
                    break;
                }
            }

            if (!found) {
                reportDisplayEl.innerHTML = `<p class="error">在報告列表中找不到名為 "${reportName}" 的檔案。</p>`;
            }
        };

        await fetchReports();
        findAndLoadReport();
    }


    function debounce(func, delay) {
        let timeoutId;
        return function(...args) {
            clearTimeout(timeoutId);
            timeoutId = setTimeout(() => {
                func.apply(this, args);
            }, delay);
        };
    }

    async function fetchReports() {
        reportListLoaderEl.classList.remove('hidden');
        reportListEl.innerHTML = '';
        try {
            const response = await fetch(`${BACKEND_URL}/api/reports`, {
                headers: { 'ngrok-skip-browser-warning': 'true' }
            });
            if (!response.ok) {
                const err = await response.json();
                throw new Error(err.detail || `HTTP Error: ${response.status}`);
            }
            const reports = await response.json();
            displayReports(reports);
        } catch (error) {
            console.error("無法獲取報告列表:", error);
            reportListEl.innerHTML = `<li class="error">無法載入報告列表: ${error.message}</li>`;
        } finally {
            reportListLoaderEl.classList.add('hidden');
        }
    }

    // 在報告列表項目中加入下載圖示
    function displayReports(reports) {
        if (reports.length === 0) {
            reportListEl.innerHTML = '<li>在指定路徑下未找到任何報告</li>';
            return;
        }
        reportListEl.innerHTML = reports.map(report => {
            // report.name 可能是 "Folder / file.docx" 或 "file.docx"
            // data-name 儲存完整路徑，用於選取
            // span 中只顯示最後的檔案名
            const displayName = report.name.includes('/') ? report.name.split('/').pop().trim() : report.name;
            const pureFilename = displayName; // 用於下載標題
            const isActive = currentReport.id === report.id ? 'active' : '';
            
            return `
            <li data-id="${report.id}" data-name="${report.name}" class="${isActive}">
                <span class="report-name" title="${report.name}">${report.name}</span>
                <a href="${BACKEND_URL}/api/download/${report.id}" 
                class="download-icon" 
                title="下載此報告: ${pureFilename}" 
                data-id="${report.id}" 
                data-filename="${pureFilename}">
                    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16">
                        <path d="M.5 9.9a.5.5 0 0 1 .5.5v2.5a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1v-2.5a.5.5 0 0 1 1 0v2.5a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2v-2.5a.5.5 0 0 1 .5-.5z"/>
                        <path d="M7.646 11.854a.5.5 0 0 0 .708 0l3-3a.5.5 0 0 0-.708-.708L8.5 10.293V1.5a.5.5 0 0 0-1 0v8.793L5.354 8.146a.5.5 0 1 0-.708.708l3 3z"/>
                    </svg>
                </a>
            </li>
        `}).join('');
    }


    function setupWebSocket() {
        const wsUrl = BACKEND_URL.replace(/^http/, 'ws') + '/ws/refine';
        ws = new WebSocket(wsUrl);
        ws.onopen = () => { console.log('WebSocket 連線成功'); updateConnectionStatus(true); };
        ws.onmessage = (event) => { handleWebSocketMessage(JSON.parse(event.data)); };
        ws.onclose = () => { console.log('WebSocket 連線斷開，5秒後嘗試重連...'); updateConnectionStatus(false); setTimeout(setupWebSocket, 5000); };
        ws.onerror = (error) => { console.error('WebSocket 錯誤:', error); updateConnectionStatus(false); ws.close(); };
    }

    function updateConnectionStatus(isConnected) {
        connectionStatusEl.className = `status-indicator ${isConnected ? 'connected' : 'disconnected'}`;
        connectionStatusEl.querySelector('.text').textContent = isConnected ? '已連接' : '未連接';
    }

    // 區分點擊報告本身和點擊下載圖示的行為
    function addEventListeners() {
        selectReportDebounced = debounce((id, name, element) => {
            selectReport(id, name, element);
        }, 300);

        reportListEl.addEventListener('click', (e) => {
            // 檢查是否點擊了下載圖示
            const downloadIcon = e.target.closest('.download-icon');
            if (downloadIcon) {
                e.preventDefault(); // ✨ 關鍵：阻止<a>標籤的預設跳轉行為
                const fileId = downloadIcon.dataset.id;
                const filename = downloadIcon.dataset.filename;
                triggerDownload(fileId, filename, downloadIcon); // ✨ 調用新的下載處理函式
                return;
            }

            if (isLoadingReport) {
                console.warn("正在加載報告，請稍候...");
                return;
            }

            const listItem = e.target.closest('li[data-id]');
            if (listItem) {
                document.querySelectorAll('#report-list li.pending').forEach(li => li.classList.remove('pending'));
                listItem.classList.add('pending');
                selectReportDebounced(listItem.dataset.id, listItem.dataset.name, listItem);
            }
        });
        reportDisplayEl.addEventListener('mouseup', handleTextSelection);
        sendInstructionBtn.addEventListener('click', sendInstruction);
        refineInstructionInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); sendInstruction(); }
        });
        saveReportBtn.addEventListener('click', saveReport);
    }

    // --- Konami Code Easter Egg ---
    const konamiSequence = [38, 38, 40, 40, 37, 39, 37, 39, 66, 65]; // ↑ ↑ ↓ ↓ ← → ← → B A
    let konamiIndex = 0;


    const ericSequence = [69, 82, 73, 67]; // E R I C 的鍵盤碼
    let ericIndex = 0;

    window.addEventListener('keydown', e => {
        // 如果使用者正在輸入框中打字，則重置所有密技序列，避免誤觸
        if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') {
            konamiIndex = 0;
            ericIndex = 0;
            return;
        }

        // --- 檢查 Konami Code ---
        if (e.keyCode === konamiSequence[konamiIndex]) {
            konamiIndex++;
            if (konamiIndex === konamiSequence.length) {
                console.log("🚀 Konami Code Activated!");
                activateSecretMode(); // 觸發第一個彩蛋
                konamiIndex = 0; // 重置索引
            }
        } else {
            konamiIndex = 0;
        }

        // --- 檢查 ERIC Code ---
        if (e.keyCode === ericSequence[ericIndex]) {
            ericIndex++;
            if (ericIndex === ericSequence.length) {
                console.log("💥 ERIC Code Activated! System Collapse Imminent!");
                triggerPageCollapse(); // 觸發第二個彩蛋
                ericIndex = 0; // 重置索引
            }
        } else {
            // 如果按錯了，只重置 ERIC 的索引
            ericIndex = 0;
        }
    });



    // 圖片保護機制：為所有圖片添加唯一 ID
    function protectImages() {
        const images = reportDisplayEl.querySelectorAll('img');
        images.forEach((img, index) => {
            if (!img.id) {
                img.id = `protected-img-${Date.now()}-${index}`;
            }
            // 確保圖片有 data-original-src 屬性作為備份
            if (!img.getAttribute('data-original-src')) {
                img.setAttribute('data-original-src', img.src);
            }
        });
        console.log(`🛡️ 保護了 ${images.length} 張圖片`);
    }


    function activateSecretMode() {
        // --- Part 1: 天崩地裂特效 ---
        document.body.classList.add('screen-shake');
        
        // --- Part 2: 自動下載 EasterEgg.jpg (在特效結束後執行) ---
        const shakeAnimationDuration = 800; // 必須與 CSS 中的動畫時間匹配
        setTimeout(() => {
            console.log("💥 特效結束，開始下載秘密檔案...");
            // 使用我們現有的下載函式，請求新的 API 端點
            triggerDownload(null, "EasterEgg.jpg", null, `${BACKEND_URL}/api/special/download-easteregg`);
            
            // 動畫結束後移除 class，以便下次觸發
            document.body.classList.remove('screen-shake');
        }, shakeAnimationDuration);

        // --- Part 3: 啟用 Turbo UI 強化模式 (您提供的程式碼) ---
        document.body.classList.toggle('turbo');
        
        // 動態打字效果的樣式注入
        const styleId = 'turbo-typing-style';
        if (!document.getElementById(styleId)) {
            const style = document.createElement('style');
            style.id = styleId;
            // 我們讓打字動畫只在 Turbo 模式啟用時生效
            style.textContent = `
                body.turbo .chat-bubble.ai .refined-content { 
                    display: inline-block; /* 讓 border-right 生效 */
                    animation: typing 2s steps(40, end), blink-caret .75s step-end infinite;
                    white-space: nowrap; 
                    overflow: hidden; 
                    border-right: .15em solid var(--secondary-color);
                    max-width: 100%;
                }
                @keyframes typing { from { width: 0 } to { width: 100% } }
                @keyframes blink-caret { from, to { border-color: transparent } 50% { border-color: var(--secondary-color); } }
            `;
            document.head.appendChild(style);
        }

        // Turbo 開關按鈕
        const buttonId = 'turbo-btn';
        if (!document.getElementById(buttonId)) {
            const btn = document.createElement('button');
            btn.id = buttonId;
            btn.className = 'btn';
            btn.style.cssText = 'position:fixed; top:20px; right:20px; z-index:1001;';
            
            const updateButtonState = () => {
                const isTurbo = document.body.classList.contains('turbo');
                btn.innerHTML = isTurbo ? '🚀&nbsp;Turbo&nbsp;On' : '🚀&nbsp;Turbo&nbsp;Off';
                if(isTurbo) {
                    btn.classList.add('btn-primary');
                } else {
                    btn.classList.remove('btn-primary');
                }
            };
            
            btn.onclick = () => {
                document.body.classList.toggle('turbo');
                updateButtonState();
            };
            
            document.body.appendChild(btn);
            updateButtonState();
        }
    }


    function triggerPageCollapse() {
        console.log("💥 執行頁面崩潰程序...");

        triggerDownload(
            null, 
            "EasterEgg2.jpg", 
            null, 
            `${BACKEND_URL}/api/special/download-easteregg2`
        );

        setTimeout(() => {
            // 創建一個全螢幕的紅色遮罩層
            const crashOverlay = document.createElement('div');
            crashOverlay.style.cssText = `
                position: fixed;
                top: 0;
                left: 0;
                width: 100vw;
                height: 100vh;
                background-color: #8B0000; /* 深紅色 */
                color: white;
                display: flex;
                justify-content: center;
                align-items: center;
                font-size: 4vw;
                font-family: 'Courier New', Courier, monospace;
                text-align: center;
                z-index: 9999;
                flex-direction: column;
                line-height: 1.5;
            `;
            crashOverlay.innerHTML = `
                <div>FATAL SYSTEM ERROR</div>
                <div><span style="background-color: white; color: #8B0000; padding: 0 10px;">ERIC.SYS CORRUPTED</span></div>
                <div>PLEASE REFRESH YOUR BROWSER.</div>
            `;
            
            // 直接替換掉整個 body 內容，造成頁面結構的完全崩潰
            document.body.innerHTML = '';
            document.body.appendChild(crashOverlay);
            document.body.style.backgroundColor = '#8B0000';

        }, 1500); // 延遲 1.5 秒，給予使用者反應時間
    }

    async function triggerDownload(fileId, filename, iconElement, directUrl = null) {
        console.log(`開始下載檔案: ${filename}`);
        const downloadUrl = directUrl ? directUrl : `${BACKEND_URL}/api/download/${fileId}`;

        if(iconElement) iconElement.classList.add('is-downloading');
        try {
            const response = await fetch(downloadUrl, {
                method: 'GET',
                headers: {
                    'ngrok-skip-browser-warning': 'true'
                }
            });

            if (!response.ok) {
                throw new Error(`下載失敗: ${response.status} ${response.statusText}`);
            }

            // 將回應內容轉換為 Blob (二進制大型物件)
            const blob = await response.blob();
            
            // 創建一個指向此 Blob 的臨時 URL
            const objectUrl = window.URL.createObjectURL(blob);
            
            // 創建一個隱形的 <a> 標籤來觸發下載
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = objectUrl;
            a.download = filename; // 設定預設存檔名稱
            
            document.body.appendChild(a);
            a.click(); // 模擬點擊以下載檔案
            
            // 清理工作
            window.URL.revokeObjectURL(objectUrl);
            document.body.removeChild(a);

        } catch (error) {
            console.error("下載過程中發生錯誤:", error);
            alert(`下載檔案 ${filename} 失敗。\n錯誤訊息: ${error.message}`);
        } finally {
            if(iconElement) iconElement.classList.remove('is-downloading');
        }
    }



    function fixImageSources() {
        const images = reportDisplayEl.querySelectorAll('img');
        let fixed = 0;
        
        images.forEach(img => {
            // 如果圖片的 src 不是 base64，可能需要修復
            if (!img.src.startsWith('data:image') && img.getAttribute('src')) {
                const originalSrc = img.getAttribute('src');
                if (originalSrc.startsWith('data:image')) {
                    img.src = originalSrc;
                    fixed++;
                }
            }
        });
        
        if (fixed > 0) {
            console.log(`修復了 ${fixed} 張圖片的 src`);
        }
    }

    // --- Core Functions ---
    async function selectReport(id, name, listItem) {
        // 防禦性檢查：如果因某些原因仍在加載，則中止
        if (isLoadingReport) return;
        
        // 如果延遲後發現要加載的還是當前已選中的報告，則中止
        if (currentReport.id === id) {
            listItem.classList.remove('pending'); // 移除「即將加載」的樣式
            return;
        }

        isLoadingReport = true; // 設定加載旗標為 true
        reportListEl.classList.add('is-loading'); // 讓報告列表變暗且不可點擊
        currentReport = { id, name };
        
        saveStatusEl.innerHTML = '';
        document.querySelectorAll('#report-list li').forEach(li => li.classList.remove('active'));
        listItem.classList.add('active'); // 正式選中
        listItem.classList.remove('pending'); // 移除「即將加載」的樣式

        welcomeScreenEl.classList.add('hidden');
        reportViewEl.classList.remove('hidden');
        refinePanelEl.classList.remove('hidden');
        reportTitleEl.textContent = name;
        reportDisplayEl.innerHTML = '<div class="loader"></div>';
        resetRefinePanel(true);


        try {
            const response = await fetch(`${BACKEND_URL}/api/report/${id}`, {
                headers: { 'ngrok-skip-browser-warning': 'true' }
            });
            if (!response.ok) throw new Error(`HTTP Error: ${response.status}`);
            const data = await response.json();
            reportDisplayEl.innerHTML = data.content;
            
            // 修復可能的圖片問題
            fixImageSources();
            protectImages(); 

            // 調試：檢查載入的圖片
            console.log("=== 報告載入後檢查 ===");
            debugCheckImages();
            
        } catch (error) {
            console.error("無法載入報告內容:", error);
            reportDisplayEl.innerHTML = '<p class="error">無法載入報告內容。</p>';
        } finally {
            isLoadingReport = false;
            reportListEl.classList.remove('is-loading');
        }
    }

    function handleTextSelection() {
        if (isApplyingChange) return;
        const selection = window.getSelection();
        const selectedText = selection.toString().trim();
        if (selectedText.length > 0) {
            currentSelection = { text: selectedText, range: selection.getRangeAt(0).cloneRange() };
            selectedTextPreviewEl.textContent = selectedText;
            selectedTextPreviewEl.classList.add('active');
            refineInstructionInput.disabled = false;
            sendInstructionBtn.disabled = false;
        }
    }

    function sendInstruction() {
        const instruction = refineInstructionInput.value.trim();
        if (!instruction || !currentSelection.text) return;
        addMessageToChat('user', instruction);
        if (ws && ws.readyState === WebSocket.OPEN) {
            ws.send(JSON.stringify({
                selection: currentSelection.text,
                instruction: instruction,
                history: chatHistory
            }));
        } else {
            addMessageToChat('ai', "連線錯誤，無法發送請求。");
        }
        refineInstructionInput.value = '';
        addMessageToChat('ai', null, true);
    }




    function debugCheckImages() {
        const content = reportDisplayEl.innerHTML;
        const tempDiv = document.createElement('div');
        tempDiv.innerHTML = content;
        
        const images = tempDiv.querySelectorAll('img');
        console.log(`📸 找到 ${images.length} 張圖片`);
        
        images.forEach((img, index) => {
            const src = img.src;
            console.log(`圖片 ${index + 1}:`, {
                src: src.substring(0, 100) + '...', // 只顯示前 100 個字符
                isBase64: src.startsWith('data:image'),
                imgElement: img
            });
        });
        
        return images.length;
    }


    async function saveReport() {
        if (!currentReport.id) return;
        if (refinedSections.length === 0) {
            saveStatusEl.innerHTML = '<span>報告未做任何修改，無需保存。</span>';
            setTimeout(() => { saveStatusEl.innerHTML = ''; }, 3000);
            return;
        }

        // 檔名修改邏輯更新
        const originalBaseName = currentReport.name.split('/').pop().trim().replace('.docx', '');
        const suggestedName = `${originalBaseName}_refined`;
        
        // 讓使用者有機會修改檔名
        let userProvidedName = prompt("若要修改檔名請輸入，或直接按「確定」以預設名稱儲存：", suggestedName);

        // 如果使用者點擊「取消」(prompt 返回 null)，我們就使用預設建議的檔名
        if (userProvidedName === null) {
            userProvidedName = suggestedName;
            console.log('使用者取消了檔名修改，使用預設名稱:', suggestedName);
        }
        
        // 如果使用者輸入了空字串，也視為取消
        if (!userProvidedName || userProvidedName.trim() === '') {
            saveStatusEl.innerHTML = '<span>已取消儲存（檔名不可為空）。</span>';
            setTimeout(() => { saveStatusEl.innerHTML = ''; }, 2000);
            return;
        }
        

        saveStatusEl.innerHTML = '';
        saveReportBtn.disabled = true;
        saveReportBtn.innerHTML = '<span class="icon">⏳</span> 保存中...';

        console.log(`使用進階保存模式，共 ${refinedSections.length} 處修改。`);
        
        try {
            const response = await fetch(`${BACKEND_URL}/api/save-advanced`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'ngrok-skip-browser-warning': 'true'
                },
                body: JSON.stringify({
                    refined_sections: refinedSections,
                    filename: currentReport.name,
                    original_file_id: currentReport.id,
                    new_filename: userProvidedName // 將最終決定的檔名傳給後端
                })
            });

            const result = await response.json();
            if (!response.ok) throw new Error(result.detail || '保存失敗');
            
            const newFile = result.file;
            console.log(`進階保存成功，新檔案 ID: ${newFile.id}, 共替換 ${result.replaced_count} 處文字。`);

            currentReport.id = newFile.id;
            currentReport.name = newFile.name;

            saveStatusEl.innerHTML = `
                <span> 保存成功! (共 ${result.replaced_count} 處修改)</span>
                <a href="#" class="download-link" data-id="${newFile.id}" data-filename="${newFile.name}">下載檔案</a> |
                <a href="${newFile.webViewLink}" target="_blank">在雲端檢視</a>
            `;

            saveStatusEl.querySelector('.download-link').addEventListener('click', (e) => {
                e.preventDefault();
                triggerDownload(e.target.dataset.id, e.target.dataset.filename, null);
            });

            refinedSections = [];
            await fetchReports();

            const newListItem = document.querySelector(`#report-list li[data-id="${newFile.id}"]`);
            if (newListItem) {
                document.querySelectorAll('#report-list li').forEach(li => li.classList.remove('active'));
                newListItem.classList.add('active');
            }

        } catch (error) {
            console.error("進階保存報告失敗:", error);
            saveStatusEl.innerHTML = `保存失敗: ${error.message}`;
        } finally {
            saveReportBtn.disabled = false;
            saveReportBtn.innerHTML = '<span class="icon">💾</span> 保存報告';
        }
    }

    // --- UI Updates & Helpers ---
    function handleWebSocketMessage(data) {
        const loadingBubble = chatHistoryEl.querySelector('.loading-bubble');
        if (loadingBubble) loadingBubble.remove();

        if (data.type === 'refinement_result') {
            const { original_instruction, refined_text } = data;
            chatHistory.push({ role: "user", content: `原始文字：${currentSelection.text}\n指示：${original_instruction}` });
            chatHistory.push({ role: "assistant", content: refined_text });
            addMessageToChat('ai', { original: original_instruction, refined: refined_text });
        } else if (data.error) {
            console.error("後端 WebSocket 錯誤:", data.error);
            addMessageToChat('ai', `發生錯誤: ${data.error}`);
        }
    }

    function addMessageToChat(sender, content, isLoading = false) {
        const bubble = document.createElement('div');
        bubble.className = `chat-bubble ${sender}`;
        if (isLoading) {
            bubble.classList.add('loading-bubble');
            bubble.innerHTML = '<div class="loader"></div>';
        } else if (sender === 'user') {
            bubble.textContent = content;
        } else if (sender === 'ai') {
            if (typeof content === 'object') {
                bubble.innerHTML = `
                    <p>根據您的指令「<strong>${content.original}</strong>」，我將文字修改為：</p>
                    <div class="refined-content">${content.refined.replace(/\n/g, '<br>')}</div>
                    <button class="btn apply-btn" data-refined-text="${encodeURIComponent(content.refined)}">應用修改</button>
                `;
                bubble.querySelector('.apply-btn').addEventListener('click', (e) => {
                    const refinedText = decodeURIComponent(e.target.dataset.refinedText);
                    applyRefinement(refinedText);
                    e.target.textContent = '已應用';
                    e.target.disabled = true;
                });
            } else {
                bubble.textContent = content;
            }
        }
        chatHistoryEl.appendChild(bubble);
        chatHistoryEl.scrollTop = chatHistoryEl.scrollHeight;
    }

    function applyRefinement(newText) {
        // 驗證選取範圍是否仍然有效
        try {
            const testRange = currentSelection.range.cloneRange();
            const testText = testRange.toString();
            if (testText !== currentSelection.text) {
                console.error('選取範圍已失效');
                alert('選取範圍已改變，請重新選取文字');
                return;
            }
        } catch (e) {
            console.error('無法驗證選取範圍:', e);
            alert('選取範圍無效，請重新選取文字');
            return;
        }
        
        if (!currentSelection.range) return;
        // 記錄這次修改
        refinedSections.push({
            original: currentSelection.text,
            refined: newText
        });
        console.log(`📝 記錄修改 #${refinedSections.length}:`, {
            original: currentSelection.text.substring(0, 50) + '...',
            refined: newText.substring(0, 50) + '...'
        });

        isApplyingChange = true;
        
        // 保存所有圖片的完整資訊
        const allImages = Array.from(reportDisplayEl.querySelectorAll('img')).map((img, index) => ({
            src: img.src,
            alt: img.alt || '',
            className: img.className || '',
            id: img.id || `img-${Date.now()}-${index}`,
            parentId: img.parentElement ? img.parentElement.id : null,
            nextSiblingId: img.nextSibling && img.nextSibling.id ? img.nextSibling.id : null,
            previousSiblingId: img.previousSibling && img.previousSibling.id ? img.previousSibling.id : null,
            outerHTML: img.outerHTML
        }));
        
        console.log(`📸 保存了 ${allImages.length} 張圖片資訊`);
        
        // 執行文字替換
        const selection = window.getSelection();
        selection.removeAllRanges();
        selection.addRange(currentSelection.range);
        
        // 創建新內容節點
        const newContentNode = document.createElement('span');
        newContentNode.innerHTML = newText.replace(/\n/g, '<br>');
        
        // 替換選中內容
        currentSelection.range.deleteContents();
        currentSelection.range.insertNode(newContentNode);
        
        // 立即檢查並修復所有圖片
        setTimeout(() => {
            allImages.forEach(imgData => {
                // 首先嘗試通過 ID 找到圖片
                let img = document.getElementById(imgData.id);
                
                // 如果找不到，嘗試其他方法
                if (!img) {
                    // 嘗試通過 src 屬性找到圖片
                    img = Array.from(reportDisplayEl.querySelectorAll('img')).find(i => 
                        i.src === imgData.src || i.getAttribute('src') === imgData.src
                    );
                }
                
                // 如果圖片存在但 src 丟失，恢復它
                if (img && (!img.src || img.src === 'about:blank' || img.src === '')) {
                    img.src = imgData.src;
                    img.alt = imgData.alt;
                    if (imgData.className) img.className = imgData.className;
                    console.log('恢復了一張圖片');
                }
                
                // 如果圖片完全不存在，嘗試重新插入
                if (!img && imgData.parentId) {
                    const parent = document.getElementById(imgData.parentId);
                    if (parent) {
                        const tempDiv = document.createElement('div');
                        tempDiv.innerHTML = imgData.outerHTML;
                        const newImg = tempDiv.firstChild;
                        
                        // 嘗試插入到原始位置
                        if (imgData.nextSiblingId) {
                            const nextSibling = document.getElementById(imgData.nextSiblingId);
                            if (nextSibling) {
                                parent.insertBefore(newImg, nextSibling);
                            } else {
                                parent.appendChild(newImg);
                            }
                        } else {
                            parent.appendChild(newImg);
                        }
                        console.log('重新插入了一張丟失的圖片');
                    }
                }
            });
            protectImages(); 

            // 再次檢查圖片
            console.log("=== 修改後檢查 ===");
            debugCheckImages();
            
            isApplyingChange = false;
        }, 100);
        
        selection.removeAllRanges();
        resetRefinePanel(false);
    }

    function resetRefinePanel(fullReset = true) {
        selectedTextPreviewEl.innerHTML = '<p class="placeholder">請在報告中用滑鼠選取文字</p>';
        selectedTextPreviewEl.classList.remove('active');
        refineInstructionInput.disabled = true;
        sendInstructionBtn.disabled = true;
        currentSelection = { range: null, text: '' };
        if (fullReset) {
            chatHistoryEl.innerHTML = '';
            chatHistory = [];
            // 如果是完全重置，也清空修改記錄
            refinedSections = [];
            console.log('重置 refine panel，清空修改記錄');
        }
    }

    // --- Start Application ---
    init();
});


