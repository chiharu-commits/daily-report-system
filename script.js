// Google Apps ScriptのウェブアプリURL（config.jsから取得）
const GOOGLE_SCRIPT_URL = getScriptUrl();

// タスクカウンター
let taskCounter = 1;

// フォームの送信処理
document.getElementById('dailyReportForm').addEventListener('submit', async function(e) {
    e.preventDefault();

    // 送信ボタンを無効化
    const submitBtn = this.querySelector('.btn-primary');
    const originalText = submitBtn.textContent;
    submitBtn.textContent = '送信中...';
    submitBtn.disabled = true;

    // 基本情報の取得
    const name = document.getElementById('name').value;
    const date = document.getElementById('date').value;

    // すべてのタスクセクションからデータを取得
    const taskSections = document.querySelectorAll('.task-section');
    const tasks = [];

    taskSections.forEach((section) => {
        const index = section.getAttribute('data-task-index');
        tasks.push({
            category: document.getElementById(`category-${index}`).value,
            taskName: document.getElementById(`taskName-${index}`).value,
            workContent: document.getElementById(`workContent-${index}`).value,
            remarks: document.getElementById(`remarks-${index}`).value,
            workHours: document.getElementById(`workHours-${index}`).value
        });
    });

    // フォームデータの構築
    const formData = {
        name: name,
        date: date,
        tasks: tasks
    };

    try {
        // Google Sheetsに送信
        const response = await fetch(GOOGLE_SCRIPT_URL, {
            method: 'POST',
            mode: 'no-cors',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(formData)
        });

        // 成功メッセージ
        showMessage(`日報を登録しました！（${tasks.length}件）`, 'success');

        // 未入力バナーをクリア
        removeMissingDatesBanner();

        // フォームのクリア
        resetForm();

    } catch (error) {
        console.error('エラー:', error);
        showMessage('登録に失敗しました。もう一度お試しください。', 'error');
    } finally {
        // ボタンを元に戻す
        submitBtn.textContent = originalText;
        submitBtn.disabled = false;
    }
});

// メッセージ表示関数
function showMessage(message, type) {
    // 既存のメッセージを削除
    const existingMessage = document.querySelector('.message');
    if (existingMessage) {
        existingMessage.remove();
    }

    // メッセージ要素を作成
    const messageDiv = document.createElement('div');
    messageDiv.className = `message message-${type}`;
    messageDiv.textContent = message;

    // フォームの前に挿入
    const form = document.getElementById('dailyReportForm');
    form.parentNode.insertBefore(messageDiv, form);

    // 3秒後に自動削除
    setTimeout(() => {
        messageDiv.remove();
    }, 3000);
}

// タスクセクションを追加する関数
function addTaskSection() {
    const tasksContainer = document.getElementById('tasksContainer');
    const taskSection = document.createElement('div');
    taskSection.className = 'task-section';
    taskSection.setAttribute('data-task-index', taskCounter);

    taskSection.innerHTML = `
        <div class="task-header">
            <h3>タスク ${taskCounter + 1}</h3>
            <button type="button" class="btn btn-remove-task" onclick="removeTaskSection(${taskCounter})">削除</button>
        </div>

        <div class="form-group">
            <label for="category-${taskCounter}">プロジェクト <span class="required">*</span></label>
            <select id="category-${taskCounter}" name="category" required>
                <option value="">選択してください</option>
            </select>
        </div>

        <div class="form-group">
            <label for="taskName-${taskCounter}">タスク名 <span class="required">*</span></label>
            <select id="taskName-${taskCounter}" name="taskName" required>
                <option value="">選択してください</option>
            </select>
        </div>

        <div class="form-group">
            <label for="workContent-${taskCounter}">作業内容 <span class="required">*</span></label>
            <textarea id="workContent-${taskCounter}" name="workContent" rows="5" required placeholder="作業内容を入力してください" lang="ja" class="ime-active"></textarea>
        </div>

        <div class="form-group">
            <label for="remarks-${taskCounter}">備考（困ったこと）</label>
            <textarea id="remarks-${taskCounter}" name="remarks" rows="3" placeholder="困ったことや特記事項があれば入力してください" lang="ja" class="ime-active"></textarea>
        </div>

        <div class="form-group">
            <label for="workHours-${taskCounter}">作業時間(h) <span class="required">*</span></label>
            <input type="number" id="workHours-${taskCounter}" name="workHours" min="0" step="0.5" required placeholder="例: 2.5">
        </div>
    `;

    tasksContainer.appendChild(taskSection);

    // 新しいタスクセクションのセレクトボックスにオプションを追加
    populateTaskSelects(taskCounter);

    taskCounter++;
    updateTaskNumbers();
}

// タスクセクションを削除する関数
function removeTaskSection(index) {
    const taskSection = document.querySelector(`.task-section[data-task-index="${index}"]`);
    if (taskSection) {
        // 最後のタスクセクションは削除しない
        const remainingSections = document.querySelectorAll('.task-section');
        if (remainingSections.length > 1) {
            taskSection.remove();
            updateTaskNumbers();
        } else {
            showMessage('最低1つのタスクが必要です。', 'error');
        }
    }
}

// タスク番号を更新する関数
function updateTaskNumbers() {
    const taskSections = document.querySelectorAll('.task-section');
    taskSections.forEach((section, index) => {
        const header = section.querySelector('.task-header h3');
        header.textContent = `タスク ${index + 1}`;
    });
}

// 特定のタスクセクションのセレクトボックスにオプションを追加
function populateTaskSelects(index) {
    const categorySelect = document.getElementById(`category-${index}`);
    CONFIG.categories.forEach(category => {
        const option = document.createElement('option');
        option.value = category;
        option.textContent = category;
        categorySelect.appendChild(option);
    });

    const taskNameSelect = document.getElementById(`taskName-${index}`);
    CONFIG.taskNames.forEach(taskName => {
        const option = document.createElement('option');
        option.value = taskName;
        option.textContent = taskName;
        taskNameSelect.appendChild(option);
    });
}

// フォームをリセットする関数
function resetForm() {
    // 基本情報のリセット
    document.getElementById('name').value = '';
    const today = new Date().toISOString().split('T')[0];
    document.getElementById('date').value = today;

    // 休みチェックボックスをリセット
    document.getElementById('isHoliday').checked = false;

    // すべてのタスクセクションを削除（最初の1つを除く）
    const taskSections = document.querySelectorAll('.task-section');
    taskSections.forEach((section, index) => {
        if (index > 0) {
            section.remove();
        }
    });

    // 最初のタスクセクションをクリア
    document.getElementById('category-0').value = '';
    document.getElementById('taskName-0').value = '';
    document.getElementById('workContent-0').value = '';
    document.getElementById('remarks-0').value = '';
    document.getElementById('workHours-0').value = '';

    // タスクセクションを有効化
    document.getElementById('category-0').disabled = false;
    document.getElementById('taskName-0').disabled = false;
    document.getElementById('workContent-0').disabled = false;
    document.getElementById('workHours-0').disabled = false;

    // タスク追加ボタンを有効化
    document.getElementById('addTaskBtn').disabled = false;

    // カウンターをリセット
    taskCounter = 1;
    updateTaskNumbers();
}

// GASから設定を読み込む
async function loadConfigFromGAS() {
    try {
        // キャッシュバスターを追加してブラウザキャッシュを回避
        const timestamp = new Date().getTime();
        const url = `${GOOGLE_SCRIPT_URL}?action=getConfig&_=${timestamp}`;
        const response = await fetch(url, {
            cache: 'no-store' // キャッシュを無効化
        });
        const result = await response.json();

        if (result.status === 'success') {
            // CONFIG変数を上書き
            CONFIG.names = result.data.names;
            CONFIG.categories = result.data.categories;
            CONFIG.taskNames = result.data.taskNames;
            console.log('✅ GASから設定を読み込みました:', result.data);
        } else {
            console.warn('⚠️ GAS設定読み込み失敗、config.jsを使用します:', result.message);
        }
    } catch (error) {
        console.warn('⚠️ GAS設定読み込みエラー、config.jsを使用します:', error);
    }
}

// config.jsを読み込んでセレクトボックスを生成
async function loadConfig() {
    try {
        // まずGASから設定を読み込む
        await loadConfigFromGAS();

        // 名前のオプションを生成
        const nameSelect = document.getElementById('name');
        CONFIG.names.forEach(name => {
            const option = document.createElement('option');
            option.value = name;
            option.textContent = name;
            nameSelect.appendChild(option);
        });

        // 最初のタスクセクションのセレクトボックスにオプションを追加
        populateTaskSelects(0);

    } catch (error) {
        console.error('設定ファイルの読み込みエラー:', error);
        showMessage('設定ファイルの読み込みに失敗しました。', 'error');
    }
}

// 休みチェックボックスの制御
function toggleHolidayMode() {
    const isHoliday = document.getElementById('isHoliday').checked;
    const taskSections = document.querySelectorAll('.task-section');
    const addTaskBtn = document.getElementById('addTaskBtn');

    if (isHoliday) {
        // 休みモード：タスクセクションを無効化
        taskSections.forEach((section, index) => {
            const categorySelect = document.getElementById(`category-${index}`);
            const taskNameSelect = document.getElementById(`taskName-${index}`);
            const workContentTextarea = document.getElementById(`workContent-${index}`);
            const remarksTextarea = document.getElementById(`remarks-${index}`);
            const workHoursInput = document.getElementById(`workHours-${index}`);

            // デフォルト値を設定
            categorySelect.value = '休み';
            taskNameSelect.value = '休暇';
            workContentTextarea.value = '休み';
            remarksTextarea.value = '';
            workHoursInput.value = '0';

            // 備考以外を無効化
            categorySelect.disabled = true;
            taskNameSelect.disabled = true;
            workContentTextarea.disabled = true;
            workHoursInput.disabled = true;
        });

        // タスク追加ボタンを無効化
        addTaskBtn.disabled = true;

        // 複数タスクがある場合は削除ボタンを無効化
        const removeButtons = document.querySelectorAll('.remove-task');
        removeButtons.forEach(btn => btn.disabled = true);

    } else {
        // 通常モード：タスクセクションを有効化
        taskSections.forEach((section, index) => {
            const categorySelect = document.getElementById(`category-${index}`);
            const taskNameSelect = document.getElementById(`taskName-${index}`);
            const workContentTextarea = document.getElementById(`workContent-${index}`);
            const remarksTextarea = document.getElementById(`remarks-${index}`);
            const workHoursInput = document.getElementById(`workHours-${index}`);

            // 有効化
            categorySelect.disabled = false;
            taskNameSelect.disabled = false;
            workContentTextarea.disabled = false;
            remarksTextarea.disabled = false;
            workHoursInput.disabled = false;

            // 値をクリア
            if (categorySelect.value === '休み') categorySelect.value = '';
            if (taskNameSelect.value === '休暇') taskNameSelect.value = '';
            if (workContentTextarea.value === '休み') workContentTextarea.value = '';
            if (workHoursInput.value === '0') workHoursInput.value = '';
        });

        // タスク追加ボタンを有効化
        addTaskBtn.disabled = false;

        // 削除ボタンを有効化
        const removeButtons = document.querySelectorAll('.remove-task');
        removeButtons.forEach(btn => btn.disabled = false);
    }
}

// 対象期間（今月1日〜昨日）の平日一覧を返す
function getWeekdaysInRange() {
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const start = new Date(today.getFullYear(), today.getMonth(), 1);
    const days = [];

    for (let d = new Date(start); d < today; d.setDate(d.getDate() + 1)) {
        const dow = d.getDay();
        if (dow !== 0 && dow !== 6) {
            const y = d.getFullYear();
            const m = String(d.getMonth() + 1).padStart(2, '0');
            const day = String(d.getDate()).padStart(2, '0');
            days.push(`${y}-${m}-${day}`);
        }
    }
    return days;
}

// GASから担当者の提出済み日付を取得
async function fetchSubmittedDates(name) {
    try {
        const timestamp = new Date().getTime();
        const url = `${GOOGLE_SCRIPT_URL}?action=getSubmittedDates&name=${encodeURIComponent(name)}&_=${timestamp}`;
        const response = await fetch(url, { cache: 'no-store' });
        const result = await response.json();
        if (result.status === 'success') {
            return result.dates; // ['2026-06-02', '2026-06-03', ...]
        }
    } catch (e) {
        console.warn('提出済み日付の取得に失敗しました:', e);
    }
    return null;
}

// 日付文字列を「6/25（水）」形式に変換
function formatDateJP(dateStr) {
    const d = new Date(dateStr + 'T00:00:00');
    const month = d.getMonth() + 1;
    const day = d.getDate();
    const dow = ['日', '月', '火', '水', '木', '金', '土'][d.getDay()];
    return `${month}/${day}（${dow}）`;
}

// 未入力バナーを表示
function showMissingDatesBanner(missingDates) {
    removeMissingDatesBanner();
    if (missingDates.length === 0) return;

    const banner = document.createElement('div');
    banner.className = 'missing-dates-banner';
    banner.id = 'missingDatesBanner';
    banner.innerHTML = `
        <span class="banner-icon">⚠️</span>
        <div class="banner-body">
            <div class="banner-title">未入力の日があります（今月分）</div>
            <div class="banner-dates">${missingDates.map(formatDateJP).join('　')}</div>
        </div>
        <button class="banner-close" onclick="removeMissingDatesBanner()" title="閉じる">×</button>
    `;

    const form = document.getElementById('dailyReportForm');
    form.parentNode.insertBefore(banner, form);
}

function removeMissingDatesBanner() {
    const existing = document.getElementById('missingDatesBanner');
    if (existing) existing.remove();
}

// 確認中バナーを表示
function showLoadingBanner() {
    removeMissingDatesBanner();
    const banner = document.createElement('div');
    banner.className = 'missing-dates-banner';
    banner.id = 'missingDatesBanner';
    banner.style.borderLeftColor = '#999';
    banner.style.background = '#f5f5f5';
    banner.innerHTML = `
        <span class="banner-icon">🔍</span>
        <div class="banner-body">
            <div class="banner-title" style="color:#555">未入力日を確認中...</div>
        </div>
    `;
    const form = document.getElementById('dailyReportForm');
    form.parentNode.insertBefore(banner, form);
}

// 名前選択時に未入力日をチェック
async function checkMissingDates(name) {
    removeMissingDatesBanner();
    if (!name) return;

    showLoadingBanner();

    const submittedDates = await fetchSubmittedDates(name);
    if (submittedDates === null) {
        removeMissingDatesBanner();
        return;
    }

    const weekdays = getWeekdaysInRange();
    const submittedSet = new Set(submittedDates);
    const missing = weekdays.filter(d => !submittedSet.has(d));

    showMissingDatesBanner(missing);
}

// ページ読み込み時の初期化
document.addEventListener('DOMContentLoaded', function() {
    // config.jsを読み込み
    loadConfig();

    // 日付フィールドに今日の日付を自動設定
    const dateInput = document.getElementById('date');
    const today = new Date().toISOString().split('T')[0];
    dateInput.value = today;

    // 休みチェックボックスのイベントリスナー
    document.getElementById('isHoliday').addEventListener('change', toggleHolidayMode);

    // タスク追加ボタンのイベントリスナー
    document.getElementById('addTaskBtn').addEventListener('click', addTaskSection);

    // 名前選択時に未入力日チェック
    document.getElementById('name').addEventListener('change', function() {
        checkMissingDates(this.value);
    });
});
