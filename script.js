// Google Apps ScriptのウェブアプリURL（デプロイ後に設定してください）
const GOOGLE_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyf8fJx5znNDEFfXxpaZE75HiWs8QSM3BKuZuPu_iJ3UWNwtjqrMpSzxAHg2sAXBLfSyA/exec';

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
            <label for="category-${taskCounter}">分類 <span class="required">*</span></label>
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
            <textarea id="workContent-${taskCounter}" name="workContent" rows="5" required placeholder="作業内容を入力してください"></textarea>
        </div>

        <div class="form-group">
            <label for="remarks-${taskCounter}">備考（困ったこと）</label>
            <textarea id="remarks-${taskCounter}" name="remarks" rows="3" placeholder="困ったことや特記事項があれば入力してください"></textarea>
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

    // カウンターをリセット
    taskCounter = 1;
    updateTaskNumbers();
}

// config.jsを読み込んでセレクトボックスを生成
function loadConfig() {
    try {
        // CONFIGオブジェクト（config.jsから読み込み）を使用

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

// ページ読み込み時の初期化
document.addEventListener('DOMContentLoaded', function() {
    // config.jsを読み込み
    loadConfig();

    // 日付フィールドに今日の日付を自動設定
    const dateInput = document.getElementById('date');
    const today = new Date().toISOString().split('T')[0];
    dateInput.value = today;

    // タスク追加ボタンのイベントリスナー
    document.getElementById('addTaskBtn').addEventListener('click', addTaskSection);
});
