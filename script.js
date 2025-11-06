document.addEventListener('DOMContentLoaded', () => {
    // ⚠️ 請將此處的 URL 替換為您 Colab 後端生成的 ngrok 網址
    const BACKEND_URL = 'https://15a77c875c78.ngrok-free.app';

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
    let currentSessionId = null;  
    let currentEditId = null; 

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
                e.preventDefault(); //阻止<a>標籤的預設跳轉行為
                const fileId = downloadIcon.dataset.id;
                const filename = downloadIcon.dataset.filename;
                triggerDownload(fileId, filename, downloadIcon); // 調用新的下載處理函式
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
        const viewHistoryBtn = document.getElementById('view-history-btn');
        if (viewHistoryBtn) {
            viewHistoryBtn.addEventListener('click', showModificationHistory);
        }
        
        // 關閉歷史面板按鈕
        const closeHistoryBtn = document.getElementById('close-history-btn');
        if (closeHistoryBtn) {
            closeHistoryBtn.addEventListener('click', closeHistoryModal);
        }
        
        // 點擊背景關閉面板
        const historyModal = document.getElementById('history-modal');
        if (historyModal) {
            historyModal.addEventListener('click', (e) => {
                if (e.target === historyModal) {
                    closeHistoryModal();
                }
            });
        }

        document.addEventListener('keydown', (e) => {
            // Ctrl+Shift+D: 查看訓練數據統計 (D = Data)
            if (e.ctrlKey && e.shiftKey && e.key === 'D') {
                e.preventDefault();
                showTrainingStats();
            }
            
            // Ctrl+Shift+A: 查看詳細分析 (A = Analysis)
            if (e.ctrlKey && e.shiftKey && e.key === 'A') {
                e.preventDefault();
                showDetailedAnalysis();
            }
            
            // Ctrl+Shift+E: 導出微調數據 (E = Export)
            if (e.ctrlKey && e.shiftKey && e.key === 'E') {
                e.preventDefault();
                exportFinetuningData();
            }
        });

        
    }


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
            console.log(`✅ 修復了 ${fixed} 張圖片的 src`);
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
                history: chatHistory,
                report_id: currentReport.id,   
                report_name: currentReport.name  
            }));
        } else {
            addMessageToChat('ai', "連線錯誤，無法發送請求。");
        }
        refineInstructionInput.value = '';
        addMessageToChat('ai', null, true);
    }

    
    function sendUserFeedback(editId, feedback, issueDescription = null) {
        /**
         * 發送用戶對AI修改的反饋
         * @param {string} editId - 編輯ID
         * @param {string} feedback - 'accepted' | 'rejected' | 'modified'
         * @param {string} issueDescription - 問題描述（可選）
         */
        if (ws && ws.readyState === WebSocket.OPEN) {
            ws.send(JSON.stringify({
                type: "user_feedback",
                edit_id: editId,
                feedback: feedback,
                issue_description: issueDescription  
            }));
            
            const logMsg = issueDescription 
                ? `已發送反饋: ${feedback} (問題: ${issueDescription}) for edit ${editId}`
                : `已發送反饋: ${feedback} for edit ${editId}`;
            console.log(logMsg);
        } else {
            console.error("WebSocket 未連接，無法發送反饋");
        }
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
        
        // 即使沒有修改也允許保存副本
        if (refinedSections.length === 0) {
            const userConfirm = confirm(
                '您尚未對此報告進行任何修改。\n\n' +
                '是否要保存一份此報告的副本?\n' +
                '(這將創建一個相同內容的新檔案)'
            );
            
            if (!userConfirm) {
                saveStatusEl.innerHTML = '<span>已取消保存。</span>';
                setTimeout(() => { saveStatusEl.innerHTML = ''; }, 2000);
                return;
            }
        }

        // 檔名修改邏輯更新
        const originalBaseName = currentReport.name.split('/').pop().trim().replace('.docx', '');
        
        // 智能處理已refined的檔名
        let suggestedName;
        if (originalBaseName.includes('_refined')) {
            // 如果已經是refined檔案，建議名稱不再添加refined
            const timestamp = new Date().toISOString().slice(0,10).replace(/-/g, '');
            suggestedName = `${originalBaseName}_v${timestamp}`;
        } else {
            suggestedName = `${originalBaseName}_refined`;
        }
        
        // 讓使用者有機會修改檔名
        let userProvidedName = prompt(
            `${refinedSections.length > 0 ? '已記錄 ' + refinedSections.length + ' 處修改。\n\n' : ''}` +
            '若要修改檔名請輸入，或直接按「確定」以預設名稱儲存：', 
            suggestedName
        );

        // 如果使用者點擊「取消」，使用預設建議的檔名
        if (userProvidedName === null) {
            userProvidedName = suggestedName;
            console.log('使用者取消了檔名修改，使用預設名稱:', suggestedName);
        }
        
        // 如果使用者輸入了空字串，也視為取消
        if (!userProvidedName || userProvidedName.trim() === '') {
            saveStatusEl.innerHTML = '<span>已取消儲存(檔名不可為空)。</span>';
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
                    new_filename: userProvidedName
                })
            });

            const result = await response.json();
            if (!response.ok) throw new Error(result.detail || '保存失敗');
            
            const newFile = result.file;
            console.log(`進階保存成功，新檔案 ID: ${newFile.id}, 共替換 ${result.replaced_count} 處文字。`);

            currentReport.id = newFile.id;
            currentReport.name = newFile.name;

            saveStatusEl.innerHTML = `
                <span> 保存成功! ${result.replaced_count > 0 ? `(共 ${result.replaced_count} 處修改)` : '(已創建副本)'}</span>
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
            saveStatusEl.innerHTML = ` 保存失敗: ${error.message}`;
        } finally {
            saveReportBtn.disabled = false;
            saveReportBtn.innerHTML = '<span class="icon">💾</span> 保存報告';
        }
    }


    // 獲取並顯示訓練數據統計
    async function showTrainingStats() {
        try {
            const response = await fetch(`${BACKEND_URL}/api/training-stats`, {
                headers: { 'ngrok-skip-browser-warning': 'true' }
            });
            
            if (!response.ok) throw new Error('無法獲取統計');
            
            const stats = await response.json();
            
            // 創建統計顯示信息
            const statsMessage = `
    ╔═══════════════════════════════════════╗
            AI 訓練數據統計報告               
    ╚═══════════════════════════════════════╝

    總文件數: ${stats.total_files}
    總樣本數: ${stats.total_samples}
    存儲位置: PaperGenerator/training_data/

    ${stats.files && stats.files.length > 0 ? `
    最近的文件:
    ${stats.files.slice(0, 5).map(f => `  • ${f.name} (${new Date(f.modifiedTime).toLocaleDateString()})`).join('\n')}
    ` : ''}

    這些數據可用於微調 AI 模型
    數據格式: JSONL (標準訓練格式)
            `;
            
            alert(statsMessage);
            console.log('訓練數據統計:', stats);
                
        } catch (error) {
            console.error('獲取統計失敗:', error);
            alert('無法獲取訓練數據統計\n\n可能原因：\n1. 後端服務未啟動\n2. 訓練數據文件夾未初始化\n3. 網絡連接問題');
        }
    }


    // 查看詳細訓練數據分析
    async function showDetailedAnalysis() {
        try {
            const response = await fetch(`${BACKEND_URL}/api/training-analysis`, {
                headers: { 'ngrok-skip-browser-warning': 'true' }
            });
            
            if (!response.ok) throw new Error('無法獲取分析');
            
            const analysis = await response.json();
            
            // 格式化編輯類型分布
            const editTypeStr = Object.entries(analysis.edit_type_distribution || {})
                .map(([type, count]) => `  • ${type}: ${count}`)
                .join('\n');
            
            const analysisMessage = `
    ╔═══════════════════════════════════════════╗
        AI 訓練數據質量分析報告               
    ╚═══════════════════════════════════════════╝

    總體統計:
    • 總樣本數: ${analysis.total_samples}
    • 平均編輯比率: ${(analysis.average_edit_ratio * 100).toFixed(2)}%
    • 多輪編輯數: ${analysis.multi_turn_edits}

    編輯類型分布:
    ${editTypeStr}

    質量指標:
    • 高質量 (用戶接受): ${analysis.quality_metrics.high_quality}
    • 需改進 (用戶拒絕): ${analysis.quality_metrics.needs_improvement}
    • 未評價: ${analysis.total_samples - analysis.quality_metrics.high_quality - analysis.quality_metrics.needs_improvement}

    建議:
    ${analysis.average_edit_ratio > 0.5 ? '✓ 編輯幅度適中，模型學習效果好' : '⚠ 編輯幅度較小，考慮收集更大改動的樣本'}
    ${analysis.quality_metrics.high_quality > analysis.total_samples * 0.7 ? '✓ 高質量樣本佔比良好' : '⚠ 建議增加高質量樣本比例'}
            `;
            
            alert(analysisMessage);
            console.log('詳細分析:', analysis);
                
        } catch (error) {
            console.error('獲取分析失敗:', error);
            alert('無法獲取訓練數據分析\n\n可能原因:\n1. 後端服務未啟動\n2. 訓練數據文件夾未初始化\n3. 網絡連接問題');
        }
    }

    // 導出微調數據
    async function exportFinetuningData() {
        const format = prompt(
            '選擇導出格式:\n\n' +
            '1. OpenAI (輸入: openai)\n' +
            '2. Anthropic/Claude (輸入: anthropic)\n\n' +
            '請輸入格式名稱:',
            'anthropic'
        );
        
        if (!format || !['openai', 'anthropic'].includes(format.toLowerCase())) {
            alert('已取消或格式無效');
            return;
        }
        
        try {
            console.log(`開始導出 ${format} 格式的微調數據...`);
            
            const downloadUrl = `${BACKEND_URL}/api/export-for-finetuning?format=${format}`;
            
            // 創建隱藏的下載鏈接
            const a = document.createElement('a');
            a.href = downloadUrl;
            a.style.display = 'none';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            
            alert(` 微調數據導出成功！\n格式: ${format.toUpperCase()}\n文件將自動下載`);
            
        } catch (error) {
            console.error('導出失敗:', error);
            alert(` 導出失敗: ${error.message}`);
        }
    }

    
    // --- UI Updates & Helpers ---
    function handleWebSocketMessage(data) {
        const loadingBubble = chatHistoryEl.querySelector('.loading-bubble');
        if (loadingBubble) loadingBubble.remove();

        if (data.type === 'refinement_result') {
            const { original_instruction, refined_text, session_id } = data;
            
            // 保存 session_id
            currentSessionId = session_id;
            
            // 生成編輯ID（從時間戳和哈希組合）
            const editId = `${Date.now()}_${Math.floor(Math.random() * 10000)}`;
            currentEditId = editId;
            
            chatHistory.push({ 
                role: "user", 
                content: `原始文字:${currentSelection.text}\n指示:${original_instruction}` 
            });
            chatHistory.push({ 
                role: "assistant", 
                content: refined_text 
            });
            
            // 傳遞 editId 給 addMessageToChat
            addMessageToChat('ai', { 
                original: original_instruction, 
                refined: refined_text,
                editId: editId  // 新增
            });
            
        } else if (data.type === 'feedback_recorded') {
            console.log(` 反饋已記錄: ${data.feedback} for ${data.edit_id}`);
            
        } else if (data.type === 'feedback_error') {
            console.error(" 反饋記錄失敗:", data.error);
            
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
                    <p>根據您的指令「<strong>${content.original}</strong>」,我將文字修改為:</p>
                    <div class="refined-content">${content.refined.replace(/\n/g, '<br>')}</div>
                    <div class="action-buttons">
                        <button class="btn apply-btn" data-refined-text="${encodeURIComponent(content.refined)}" data-edit-id="${content.editId}">
                            ✓ 應用修改
                        </button>
                        <button class="btn feedback-btn feedback-good" data-edit-id="${content.editId}" title="這個修改很好">
                            👍
                        </button>
                        <button class="btn feedback-btn feedback-bad" data-edit-id="${content.editId}" title="這個修改需要改進">
                            👎
                        </button>
                    </div>
                `;
                
                // 應用修改按鈕
                bubble.querySelector('.apply-btn').addEventListener('click', (e) => {
                    const refinedText = decodeURIComponent(e.target.dataset.refinedText);
                    const editId = e.target.dataset.editId;
                    
                    applyRefinement(refinedText);
                    
                    // 自動發送 "accepted" 反饋
                    sendUserFeedback(editId, 'accepted');
                    
                    e.target.textContent = '✓ 已應用';
                    e.target.disabled = true;
                    
                    // 禁用反饋按鈕
                    bubble.querySelectorAll('.feedback-btn').forEach(btn => btn.disabled = true);
                });
                
                // 好評按鈕
                bubble.querySelector('.feedback-good').addEventListener('click', (e) => {
                    const editId = e.target.dataset.editId;
                    sendUserFeedback(editId, 'accepted');
                    
                    e.target.textContent = '✓ 已標記為好';
                    e.target.disabled = true;
                    bubble.querySelector('.feedback-bad').disabled = true;
                });
                
                // 差評按鈕
                bubble.querySelector('.feedback-bad').addEventListener('click', async (e) => {
                    const editId = e.target.dataset.editId;
                    const originalRefinedText = decodeURIComponent(
                        bubble.querySelector('.apply-btn').dataset.refinedText
                    );
                    
                    const issueDescription = prompt(
                        '這個修改有什麼問題？請簡短說明，AI 會根據您的反饋重新生成。\n\n' +
                        '例如：\n' +
                        '• 太簡短，需要更詳細\n' +
                        '• 不夠專業\n' +
                        '• 偏離原意\n' +
                        '• 語氣不對',
                        ''
                    );
                    
                    // 用戶取消或輸入空白
                    if (!issueDescription || issueDescription.trim() === '') {
                        // 只記錄差評，不重新生成
                        sendUserFeedback(editId, 'rejected');
                        e.target.textContent = '✓ 已標記為差';
                        e.target.disabled = true;
                        bubble.querySelector('.feedback-good').disabled = true;
                        return;
                    }
                    
                    // 只禁用按鈕，不在原氣泡內添加任何東西
                    e.target.textContent = '✓ 已反饋';
                    e.target.disabled = true;
                    bubble.querySelector('.feedback-good').disabled = true;
                    bubble.querySelector('.apply-btn').disabled = true;
                    
                    // 記錄差評和問題描述
                    sendUserFeedback(editId, 'rejected', issueDescription);
                    
                    // 創建新的獨立用戶消息氣泡
                    addMessageToChat('user', `這個修改有問題：${issueDescription}\n\n請重新修改。`);
                    
                    // 構建新指令
                    const newInstruction = `之前的修改有問題：${issueDescription}\n\n請根據這個反饋重新修改原始文字。`;
                    
                    // 發送重新生成請求
                    if (ws && ws.readyState === WebSocket.OPEN) {
                        ws.send(JSON.stringify({
                            selection: currentSelection.text,
                            instruction: newInstruction,
                            history: chatHistory,
                            report_id: currentReport.id,
                            report_name: currentReport.name,
                            is_regeneration: true,
                            previous_attempt: originalRefinedText
                        }));
                        
                        // 顯示新的獨立載入氣泡
                        addMessageToChat('ai', null, true);
                    } else {
                        addMessageToChat('ai', '連線錯誤，無法重新生成。');
                    }
                });
                
            } else {
                bubble.textContent = content;
            }
        }
        chatHistoryEl.appendChild(bubble);
        chatHistoryEl.scrollTop = chatHistoryEl.scrollHeight;
    }

    function applyRefinement(newText) {
        if (!currentSelection.text) {
            alert('沒有選取的文字');
            return;
        }
        
        // 嘗試驗證現有 range
        let rangeIsValid = false;
        try {
            const testRange = currentSelection.range.cloneRange();
            const testText = testRange.toString();
            if (testText === currentSelection.text) {
                rangeIsValid = true;
            }
        } catch (e) {
            console.warn('原始 range 已失效，將重新尋找文字位置');
        }
        
        // 如果 range 失效，重新尋找文字位置
        if (!rangeIsValid) {
            const found = findAndSelectText(currentSelection.text);
            if (!found) {
                alert('無法在文檔中找到原始文字，可能已被修改。請重新選取文字。');
                return;
            }
            console.log('已重新定位原始文字');
        }
        
        // 記錄這次修改
        refinedSections.push({
            original: currentSelection.text,
            refined: newText
        });
        console.log(`記錄修改 #${refinedSections.length}:`, {
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
        
        console.log(`保存了 ${allImages.length} 張圖片資訊`);
        
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

    /**
     * 在文檔中尋找並選取指定文字
     * @param {string} text - 要尋找的文字
     * @returns {boolean} - 是否找到並成功選取
     */
    function findAndSelectText(text) {
        try {
            // 使用 window.find() API (適用於大多數瀏覽器)
            const found = window.find(text, false, false, true, false, true, false);
            
            if (found) {
                const selection = window.getSelection();
                if (selection.rangeCount > 0) {
                    currentSelection.range = selection.getRangeAt(0).cloneRange();
                    return true;
                }
            }
            
            // 如果 window.find 失敗，手動搜索
            return findTextInElement(reportDisplayEl, text);
            
        } catch (e) {
            console.error('尋找文字時出錯:', e);
            return false;
        }
    }

    /**
     * 在指定元素中遞歸尋找文字並創建選取範圍
     * @param {HTMLElement} element - 要搜索的元素
     * @param {string} searchText - 要尋找的文字
     * @returns {boolean} - 是否找到
     */
    function findTextInElement(element, searchText) {
        const walker = document.createTreeWalker(
            element,
            NodeFilter.SHOW_TEXT,
            null,
            false
        );
        
        let node;
        let combinedText = '';
        let textNodes = [];
        
        // 收集所有文字節點
        while (node = walker.nextNode()) {
            textNodes.push({
                node: node,
                startIndex: combinedText.length,
                text: node.textContent
            });
            combinedText += node.textContent;
        }
        
        // 在組合的文字中尋找目標文字
        const index = combinedText.indexOf(searchText);
        if (index === -1) {
            return false;
        }
        
        // 找到包含目標文字的節點
        const endIndex = index + searchText.length;
        let startNode = null;
        let startOffset = 0;
        let endNode = null;
        let endOffset = 0;
        
        for (let i = 0; i < textNodes.length; i++) {
            const textNode = textNodes[i];
            const nodeEndIndex = textNode.startIndex + textNode.text.length;
            
            // 找到起始節點
            if (!startNode && index >= textNode.startIndex && index < nodeEndIndex) {
                startNode = textNode.node;
                startOffset = index - textNode.startIndex;
            }
            
            // 找到結束節點
            if (endIndex > textNode.startIndex && endIndex <= nodeEndIndex) {
                endNode = textNode.node;
                endOffset = endIndex - textNode.startIndex;
                break;
            }
        }
        
        // 創建新的選取範圍
        if (startNode && endNode) {
            const range = document.createRange();
            range.setStart(startNode, startOffset);
            range.setEnd(endNode, endOffset);
            
            currentSelection.range = range;
            
            // 視覺化選取（可選）
            const selection = window.getSelection();
            selection.removeAllRanges();
            selection.addRange(range);
            
            return true;
        }
        
        return false;
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
    
    // ==================== 修改歷史功能 ====================

async function showModificationHistory() {
    if (!currentReport.id) {
        alert('請先選擇一份報告');
        return;
    }
    
    const historyModal = document.getElementById('history-modal');
    const historyLoading = document.getElementById('history-loading');
    const historyStats = document.getElementById('history-stats');
    const historyList = document.getElementById('history-list');
    
    // 顯示模態框和載入動畫
    historyModal.classList.remove('hidden');
    historyLoading.classList.remove('hidden');
    historyStats.innerHTML = '';
    historyList.innerHTML = '';
    
    try {
        const response = await fetch(`${BACKEND_URL}/api/report-history/${currentReport.id}`, {
            headers: { 'ngrok-skip-browser-warning': 'true' }
        });
        
        if (!response.ok) {
            throw new Error('無法獲取修改歷史');
        }
        
        const data = await response.json();
        
        console.log('📜 修改歷史數據:', data);
        
        // 隱藏載入動畫
        historyLoading.classList.add('hidden');
        
        // 顯示統計信息
        displayHistoryStats(data, historyStats);
        
        // 顯示修改記錄
        displayHistoryList(data.modifications, historyList);
        
    } catch (error) {
        console.error('獲取修改歷史失敗:', error);
        historyLoading.classList.add('hidden');
        historyList.innerHTML = `
            <div class="empty-state">
                <div class="empty-state-icon">⚠️</div>
                <p class="empty-state-text">無法載入修改歷史：${error.message}</p>
            </div>
        `;
    }
}

function displayHistoryStats(data, container) {
    if (data.total_modifications === 0) {
        container.innerHTML = `
            <div class="empty-state">
                <div class="empty-state-icon">📭</div>
                <p class="empty-state-text">此報告尚無修改記錄</p>
            </div>
        `;
        return;
    }
    
    container.innerHTML = `
        <div class="stat-item">
            <span class="stat-value">${data.total_modifications}</span>
            <span class="stat-label">總修改次數</span>
        </div>
        <div class="stat-item">
            <span class="stat-value">${data.unique_segments}</span>
            <span class="stat-label">修改段落數</span>
        </div>
        <div class="stat-item">
            <span class="stat-value">${(data.total_modifications / data.unique_segments).toFixed(1)}</span>
            <span class="stat-label">平均修改次數</span>
        </div>
    `;
}

function displayHistoryList(modifications, container) {
    if (Object.keys(modifications).length === 0) {
        container.innerHTML = `
            <div class="empty-state">
                <div class="empty-state-icon">📭</div>
                <p class="empty-state-text">沒有找到任何修改記錄</p>
            </div>
        `;
        return;
    }
    
    let html = '';
    let segmentIndex = 0;
    
    for (const [originalText, versions] of Object.entries(modifications)) {
        segmentIndex++;
        
        html += `
            <div class="history-segment">
                <div class="segment-original">
                    <span class="segment-original-label">原始段落 #${segmentIndex}</span>
                    <div class="segment-original-text">${escapeHtml(originalText.substring(0, 200))}${originalText.length > 200 ? '...' : ''}</div>
                </div>
                
                <div class="segment-versions">
                    <div class="version-header">
                        <span class="version-title">修改版本</span>
                        <span class="version-count">${versions.length} 個版本</span>
                    </div>
                    ${versions.map((version, idx) => createVersionHTML(version, idx + 1, versions.length)).join('')}
                </div>
            </div>
        `;
    }
    
    container.innerHTML = html;
}

function createVersionHTML(version, versionNum, totalVersions) {
    const timestamp = new Date(version.timestamp);
    const formattedDate = timestamp.toLocaleString('zh-TW', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit'
    });
    
    const editTypeBadge = version.edit_type ? 
        `<span class="edit-type-badge edit-type-${version.edit_type}">${getEditTypeLabel(version.edit_type)}</span>` : '';
    
    const feedbackBadge = version.user_feedback ? 
        `<span class="feedback-badge feedback-${version.user_feedback}">
            ${version.user_feedback === 'accepted' ? '已接受' : version.user_feedback === 'rejected' ? '已拒絕' : '已修改'}
        </span>` : '';
    
    const stats = version.edit_statistics || {};
    const editRatio = stats.edit_ratio ? (stats.edit_ratio * 100).toFixed(1) : '0';
    
    return `
        <div class="version-item">
            <div class="version-meta">
                <div class="meta-item">
                    <span class="meta-label">版本:</span>
                    <span class="meta-value">V${versionNum}/${totalVersions}</span>
                </div>
                <div class="meta-item">
                    <span class="meta-label">時間:</span>
                    <span class="meta-value">${formattedDate}</span>
                </div>
                <div class="meta-item">
                    <span class="meta-label">修改幅度:</span>
                    <span class="meta-value">${editRatio}%</span>
                </div>
                ${editTypeBadge ? `<div class="meta-item">${editTypeBadge}</div>` : ''}
                ${feedbackBadge ? `<div class="meta-item">${feedbackBadge}</div>` : ''}
            </div>
            
            ${version.instruction ? `
                <div class="version-instruction" style="margin-bottom: ${version.refined ? 'var(--space-sm)' : '0'}; padding: var(--space-sm); background: rgba(255, 255, 0, 0.05); border-radius: 4px; font-size: 0.85rem; color: var(--gray-light);">
                    <strong style="color: var(--glow-accent);">用戶指示:</strong> ${escapeHtml(version.instruction)}
                </div>
            ` : ''}
            
            <div class="version-refined">
                <div class="version-refined-text">${escapeHtml(version.refined)}</div>
            </div>
        </div>
    `;
}

function getEditTypeLabel(type) {
    const labels = {
        'expansion': '擴展',
        'simplification': '簡化',
        'rewriting': '改寫',
        'correction': '修正',
        'style_formal': '正式化',
        'style_casual': '口語化',
        'refinement': '優化'
    };
    return labels[type] || type;
}

function closeHistoryModal() {
    const historyModal = document.getElementById('history-modal');
    historyModal.classList.add('hidden');
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

    // --- Start Application ---
    init();
});

