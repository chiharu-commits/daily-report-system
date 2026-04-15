// 設定管理画面のJavaScript

// Google Apps ScriptのウェブアプリURL（config.jsから取得）
const GOOGLE_SCRIPT_URL = getScriptUrl();

// 現在の設定データ
let currentConfig = {
  names: [],
  categories: [],
  taskNames: []
};

// ページ読み込み時の初期化
document.addEventListener('DOMContentLoaded', async () => {
  await loadConfig();
  setupEventListeners();
  setupToggleListeners();
  loadCollapsedStates();
});

// イベントリスナーの設定
function setupEventListeners() {
  // 追加ボタン
  document.getElementById('btn-add-name').addEventListener('click', () => addItem('name'));
  document.getElementById('btn-add-category').addEventListener('click', () => addItem('category'));
  document.getElementById('btn-add-task-name').addEventListener('click', () => addItem('taskName'));

  // Enterキーで追加
  document.getElementById('new-name').addEventListener('keypress', (e) => {
    if (e.key === 'Enter') addItem('name');
  });
  document.getElementById('new-category').addEventListener('keypress', (e) => {
    if (e.key === 'Enter') addItem('category');
  });
  document.getElementById('new-task-name').addEventListener('keypress', (e) => {
    if (e.key === 'Enter') addItem('taskName');
  });

  // 検索機能
  document.getElementById('search-name').addEventListener('input', (e) => {
    filterList('name-list', e.target.value);
    updateCount('name');
  });
  document.getElementById('search-category').addEventListener('input', (e) => {
    filterList('category-list', e.target.value);
    updateCount('category');
  });
  document.getElementById('search-task-name').addEventListener('input', (e) => {
    filterList('task-name-list', e.target.value);
    updateCount('taskName');
  });

  // 保存ボタン（通常 + フローティング）
  document.getElementById('btn-save').addEventListener('click', saveConfig);
  document.getElementById('btn-floating-save').addEventListener('click', saveConfig);
}

// 設定をGASから読み込む
async function loadConfig() {
  showLoading();

  try {
    // キャッシュバスターを追加してブラウザキャッシュを回避
    const timestamp = new Date().getTime();
    const url = `${GOOGLE_SCRIPT_URL}?action=getConfig&_=${timestamp}`;
    const response = await fetch(url, {
      cache: 'no-store' // キャッシュを無効化
    });
    const result = await response.json();

    if (result.status === 'success') {
      currentConfig = result.data;
      displayConfig();
      showToast('設定を読み込みました', 'success');
    } else {
      throw new Error(result.message || '設定の読み込みに失敗しました');
    }

  } catch (error) {
    console.error('Error:', error);
    showToast('設定の読み込みに失敗しました: ' + error.message, 'error');
  } finally {
    hideLoading();
  }
}

// 設定を画面に表示
function displayConfig() {
  // 名前リスト
  const nameList = document.getElementById('name-list');
  nameList.innerHTML = '';
  currentConfig.names.forEach((name, index) => {
    const li = createListItem(name, 'name', index);
    nameList.appendChild(li);
  });

  // 分類リスト
  const categoryList = document.getElementById('category-list');
  categoryList.innerHTML = '';
  currentConfig.categories.forEach((category, index) => {
    const li = createListItem(category, 'category', index);
    categoryList.appendChild(li);
  });

  // タスク名リスト
  const taskNameList = document.getElementById('task-name-list');
  taskNameList.innerHTML = '';
  currentConfig.taskNames.forEach((taskName, index) => {
    const li = createListItem(taskName, 'taskName', index);
    taskNameList.appendChild(li);
  });

  // カウント更新
  updateCount('name');
  updateCount('category');
  updateCount('taskName');

  // 検索フィールドをクリア
  document.getElementById('search-name').value = '';
  document.getElementById('search-category').value = '';
  document.getElementById('search-task-name').value = '';
}

// リストアイテムを作成
function createListItem(text, type, index) {
  const li = document.createElement('li');

  const span = document.createElement('span');
  span.className = 'item-text';
  span.textContent = text;

  const deleteBtn = document.createElement('button');
  deleteBtn.className = 'btn-delete';
  deleteBtn.textContent = '削除';
  deleteBtn.addEventListener('click', () => deleteItem(type, index));

  li.appendChild(span);
  li.appendChild(deleteBtn);

  return li;
}

// アイテムを追加
function addItem(type) {
  let inputId, arrayKey;

  if (type === 'name') {
    inputId = 'new-name';
    arrayKey = 'names';
  } else if (type === 'category') {
    inputId = 'new-category';
    arrayKey = 'categories';
  } else if (type === 'taskName') {
    inputId = 'new-task-name';
    arrayKey = 'taskNames';
  }

  const input = document.getElementById(inputId);
  const value = input.value.trim();

  if (!value) {
    showToast('値を入力してください', 'warning');
    return;
  }

  // 重複チェック
  if (currentConfig[arrayKey].includes(value)) {
    showToast('既に存在します', 'warning');
    return;
  }

  // 追加
  currentConfig[arrayKey].push(value);
  input.value = '';
  displayConfig();
  showToast('追加しました（保存ボタンを押してください）', 'success');
}

// アイテムを削除
function deleteItem(type, index) {
  let arrayKey, itemName;

  if (type === 'name') {
    arrayKey = 'names';
    itemName = currentConfig.names[index];
  } else if (type === 'category') {
    arrayKey = 'categories';
    itemName = currentConfig.categories[index];
  } else if (type === 'taskName') {
    arrayKey = 'taskNames';
    itemName = currentConfig.taskNames[index];
  }

  if (confirm(`「${itemName}」を削除してもよろしいですか？`)) {
    currentConfig[arrayKey].splice(index, 1);
    displayConfig();
    showToast('削除しました（保存ボタンを押してください）', 'success');
  }
}

// 設定をGASに保存
async function saveConfig() {
  if (!confirm('設定を保存してもよろしいですか？\nすべての画面に反映されます。')) {
    return;
  }

  showLoading();

  try {
    console.log('保存する設定データ:', currentConfig);
    console.log('送信先URL:', GOOGLE_SCRIPT_URL);

    const response = await fetch(GOOGLE_SCRIPT_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: `action=updateConfig&data=${encodeURIComponent(JSON.stringify(currentConfig))}`
    });

    console.log('レスポンスステータス:', response.status);
    const result = await response.json();
    console.log('レスポンス結果:', result);

    if (result.status === 'success') {
      showToast('設定を保存しました', 'success');
    } else {
      throw new Error(result.message || '保存に失敗しました');
    }

  } catch (error) {
    console.error('Error:', error);
    showToast('保存に失敗しました: ' + error.message, 'error');
  } finally {
    hideLoading();
  }
}

// ローディング表示
function showLoading() {
  document.getElementById('loading').classList.add('show');
}

// ローディング非表示
function hideLoading() {
  document.getElementById('loading').classList.remove('show');
}

// トースト通知
function showToast(message, type = 'success') {
  const toast = document.getElementById('toast');
  toast.textContent = message;
  toast.className = `toast show ${type}`;

  setTimeout(() => {
    toast.classList.remove('show');
  }, 3000);
}

// リストをフィルタリング
function filterList(listId, searchText) {
  const list = document.getElementById(listId);
  const items = list.querySelectorAll('li:not(.no-results)');
  const searchLower = searchText.toLowerCase().trim();

  let visibleCount = 0;

  items.forEach(item => {
    const text = item.querySelector('.item-text').textContent.toLowerCase();
    if (text.includes(searchLower)) {
      item.classList.remove('hidden');
      visibleCount++;
    } else {
      item.classList.add('hidden');
    }
  });

  // 「該当なし」メッセージの表示/非表示
  let noResults = list.querySelector('.no-results');
  if (visibleCount === 0 && items.length > 0) {
    if (!noResults) {
      noResults = document.createElement('li');
      noResults.className = 'no-results';
      noResults.textContent = '該当する項目が見つかりません';
      list.appendChild(noResults);
    }
  } else {
    if (noResults) {
      noResults.remove();
    }
  }
}

// カウント表示を更新
function updateCount(type) {
  let countId, listId, total;

  if (type === 'name') {
    countId = 'name-count';
    listId = 'name-list';
    total = currentConfig.names.length;
  } else if (type === 'category') {
    countId = 'category-count';
    listId = 'category-list';
    total = currentConfig.categories.length;
  } else if (type === 'taskName') {
    countId = 'task-name-count';
    listId = 'task-name-list';
    total = currentConfig.taskNames.length;
  }

  const list = document.getElementById(listId);
  const items = list.querySelectorAll('li:not(.no-results)');
  const visible = Array.from(items).filter(item => !item.classList.contains('hidden')).length;

  const countElement = document.getElementById(countId);
  if (visible === total) {
    countElement.textContent = `(全${total}件)`;
  } else {
    countElement.textContent = `(${visible}/${total}件表示)`;
  }
}

// 折りたたみ機能のイベントリスナー設定
function setupToggleListeners() {
  const toggleButtons = document.querySelectorAll('[data-toggle]');

  toggleButtons.forEach(button => {
    button.addEventListener('click', () => {
      const sectionType = button.getAttribute('data-toggle');
      const section = document.querySelector(`[data-section="${sectionType}"]`);

      section.classList.toggle('collapsed');
      saveCollapsedState(sectionType, section.classList.contains('collapsed'));
    });
  });
}

// 折りたたみ状態を保存
function saveCollapsedState(sectionType, isCollapsed) {
  const states = JSON.parse(localStorage.getItem('configManagerCollapsed') || '{}');
  states[sectionType] = isCollapsed;
  localStorage.setItem('configManagerCollapsed', JSON.stringify(states));
}

// 折りたたみ状態を読み込み
function loadCollapsedStates() {
  const states = JSON.parse(localStorage.getItem('configManagerCollapsed') || '{}');

  Object.keys(states).forEach(sectionType => {
    if (states[sectionType]) {
      const section = document.querySelector(`[data-section="${sectionType}"]`);
      if (section) {
        section.classList.add('collapsed');
      }
    }
  });
}
