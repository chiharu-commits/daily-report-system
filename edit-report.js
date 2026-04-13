// 日報修正画面のJavaScript

// グローバル変数
let allReportsData = []; // 検索結果を保持
let currentEditIndex = null; // 編集中のデータインデックス
let currentDeleteIndex = null; // 削除対象のデータインデックス

// 初期化
document.addEventListener('DOMContentLoaded', () => {
  initializeSelects();
  initializeEventListeners();
  setDefaultDates();
});

// セレクトボックスの初期化
function initializeSelects() {
  // 名前
  const searchNameSelect = document.getElementById('search-name');
  const addNameSelect = document.getElementById('add-name');
  const editNameSelect = document.getElementById('edit-name');
  CONFIG.names.forEach(name => {
    searchNameSelect.add(new Option(name, name));
    addNameSelect.add(new Option(name, name));
    editNameSelect.add(new Option(name, name));
  });

  // カテゴリ
  const addCategorySelect = document.getElementById('add-category');
  const editCategorySelect = document.getElementById('edit-category');
  CONFIG.categories.forEach(category => {
    addCategorySelect.add(new Option(category, category));
    editCategorySelect.add(new Option(category, category));
  });

  // タスク名
  const addTaskNameSelect = document.getElementById('add-task-name');
  const editTaskNameSelect = document.getElementById('edit-task-name');
  CONFIG.taskNames.forEach(taskName => {
    addTaskNameSelect.add(new Option(taskName, taskName));
    editTaskNameSelect.add(new Option(taskName, taskName));
  });
}

// イベントリスナーの初期化
function initializeEventListeners() {
  // 検索タイプの切り替え
  document.querySelectorAll('input[name="search-type"]').forEach(radio => {
    radio.addEventListener('change', (e) => {
      toggleSearchType(e.target.value);
    });
  });

  // クイック選択ボタン
  document.querySelectorAll('.btn-quick').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const period = e.target.dataset.period;
      setQuickPeriod(period);
    });
  });

  // 検索ボタン
  document.getElementById('btn-search').addEventListener('click', searchReports);

  // リセットボタン
  document.getElementById('btn-reset').addEventListener('click', resetSearch);

  // 新規登録ボタン
  document.getElementById('btn-add-new').addEventListener('click', openAddModal);

  // Excel出力ボタン
  document.getElementById('btn-export-excel').addEventListener('click', exportToExcel);

  // 新規登録フォーム送信
  document.getElementById('add-form').addEventListener('submit', (e) => {
    e.preventDefault();
    addReport();
  });

  // 編集フォーム送信
  document.getElementById('edit-form').addEventListener('submit', (e) => {
    e.preventDefault();
    updateReport();
  });

  // モーダル関連
  document.getElementById('btn-close-add-modal').addEventListener('click', closeAddModal);
  document.getElementById('btn-cancel-add').addEventListener('click', closeAddModal);
  document.getElementById('btn-close-modal').addEventListener('click', closeEditModal);
  document.getElementById('btn-cancel-edit').addEventListener('click', closeEditModal);
  document.getElementById('btn-close-delete-modal').addEventListener('click', closeDeleteModal);
  document.getElementById('btn-cancel-delete').addEventListener('click', closeDeleteModal);
  document.getElementById('btn-confirm-delete').addEventListener('click', deleteReport);

  // モーダル外クリックで閉じる
  document.getElementById('add-modal').addEventListener('click', (e) => {
    if (e.target.id === 'add-modal') closeAddModal();
  });
  document.getElementById('edit-modal').addEventListener('click', (e) => {
    if (e.target.id === 'edit-modal') closeEditModal();
  });
  document.getElementById('delete-modal').addEventListener('click', (e) => {
    if (e.target.id === 'delete-modal') closeDeleteModal();
  });
}

// デフォルト日付を設定（今月）
function setDefaultDates() {
  const today = new Date();
  const firstDay = new Date(today.getFullYear(), today.getMonth(), 1);
  const lastDay = new Date(today.getFullYear(), today.getMonth() + 1, 0);

  document.getElementById('start-date').valueAsDate = firstDay;
  document.getElementById('end-date').valueAsDate = lastDay;
  document.getElementById('single-date').valueAsDate = today;
}

// 検索タイプの切り替え
function toggleSearchType(type) {
  const periodSearch = document.getElementById('period-search');
  const singleSearch = document.getElementById('single-search');

  if (type === 'period') {
    periodSearch.style.display = 'block';
    singleSearch.style.display = 'none';
  } else {
    periodSearch.style.display = 'none';
    singleSearch.style.display = 'block';
  }
}

// クイック期間選択
function setQuickPeriod(period) {
  const today = new Date();
  let startDate, endDate, singleDate;

  switch (period) {
    case 'today':
      // 今日
      singleDate = new Date(today);
      document.getElementById('single-date').valueAsDate = singleDate;
      return;

    case 'yesterday':
      // 昨日
      singleDate = new Date(today);
      singleDate.setDate(today.getDate() - 1);
      document.getElementById('single-date').valueAsDate = singleDate;
      return;

    case 'thisWeek':
      // 今週（月曜日〜日曜日）
      const dayOfWeek = today.getDay();
      const diff = dayOfWeek === 0 ? -6 : 1 - dayOfWeek; // 日曜日の場合は-6、それ以外は月曜日までの差分
      startDate = new Date(today);
      startDate.setDate(today.getDate() + diff);
      endDate = new Date(startDate);
      endDate.setDate(startDate.getDate() + 6);
      break;

    case 'lastWeek':
      // 先週（月曜日〜日曜日）
      const lastWeekDay = today.getDay();
      const lastWeekDiff = lastWeekDay === 0 ? -6 : 1 - lastWeekDay;
      startDate = new Date(today);
      startDate.setDate(today.getDate() + lastWeekDiff - 7);
      endDate = new Date(startDate);
      endDate.setDate(startDate.getDate() + 6);
      break;

    case 'thisMonth':
      // 今月
      startDate = new Date(today.getFullYear(), today.getMonth(), 1);
      endDate = new Date(today.getFullYear(), today.getMonth() + 1, 0);
      break;

    case 'lastMonth':
      // 先月
      startDate = new Date(today.getFullYear(), today.getMonth() - 1, 1);
      endDate = new Date(today.getFullYear(), today.getMonth(), 0);
      break;
  }

  document.getElementById('start-date').valueAsDate = startDate;
  document.getElementById('end-date').valueAsDate = endDate;
}

// 検索実行
async function searchReports() {
  const name = document.getElementById('search-name').value;
  const searchType = document.querySelector('input[name="search-type"]:checked').value;

  // バリデーション
  if (!name) {
    showToast('名前を選択してください', 'warning');
    return;
  }

  let startDate, endDate;

  if (searchType === 'period') {
    // 期間指定
    startDate = document.getElementById('start-date').value;
    endDate = document.getElementById('end-date').value;

    if (!startDate || !endDate) {
      showToast('期間を入力してください', 'warning');
      return;
    }

    if (new Date(startDate) > new Date(endDate)) {
      showToast('開始日は終了日より前の日付を指定してください', 'warning');
      return;
    }
  } else {
    // 日付指定
    const singleDate = document.getElementById('single-date').value;

    if (!singleDate) {
      showToast('日付を入力してください', 'warning');
      return;
    }

    startDate = singleDate;
    endDate = singleDate;
  }

  // データ取得
  showLoading();

  try {
    const scriptUrl = getScriptUrl();
    const response = await fetch(scriptUrl);

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }

    const data = await response.json();

    // フィルタリング
    const filteredData = data.filter(item => {
      const itemDate = formatDateToYYYYMMDD(new Date(item.date));
      return item.name === name &&
             itemDate >= startDate &&
             itemDate <= endDate;
    });

    // 日付でソート（昇順：古い順→新しい順）
    filteredData.sort((a, b) => {
      const dateA = new Date(a.date);
      const dateB = new Date(b.date);
      return dateA - dateB;
    });

    allReportsData = filteredData;
    displayResults(filteredData);

  } catch (error) {
    console.error('Error:', error);
    showToast('データの取得に失敗しました', 'error');
  } finally {
    hideLoading();
  }
}

// 検索結果を表示
function displayResults(data) {
  const container = document.getElementById('results-container');
  const countBadge = document.getElementById('result-count');
  const exportBtn = document.getElementById('btn-export-excel');

  countBadge.textContent = `${data.length}件`;

  if (data.length === 0) {
    container.innerHTML = '<p class="no-data">該当するデータがありません</p>';
    exportBtn.disabled = true;
    return;
  }

  // Excel出力ボタンを有効化
  exportBtn.disabled = false;

  // テーブル作成
  let html = `
    <table class="results-table">
      <thead>
        <tr>
          <th>日付</th>
          <th>分類</th>
          <th>タスク名</th>
          <th>作業内容</th>
          <th>時間</th>
          <th>操作</th>
        </tr>
      </thead>
      <tbody>
  `;

  data.forEach((item, index) => {
    const displayDate = formatDateToJapanese(new Date(item.date));
    html += `
      <tr>
        <td>${displayDate}</td>
        <td>${item.category}</td>
        <td>${item.taskName}</td>
        <td>${item.workContent}</td>
        <td>${item.workHours}</td>
        <td>
          <div class="action-buttons">
            <button class="btn-edit" onclick="openEditModal(${index})">編集</button>
            <button class="btn-delete" onclick="openDeleteModal(${index})">削除</button>
          </div>
        </td>
      </tr>
    `;
  });

  html += `
      </tbody>
    </table>
  `;

  container.innerHTML = html;
}

// 新規登録モーダルを開く
function openAddModal() {
  // フォームをリセット
  document.getElementById('add-form').reset();

  // 今日の日付を設定
  const today = new Date();
  document.getElementById('add-date').valueAsDate = today;

  // モーダルを表示
  document.getElementById('add-modal').classList.add('show');
}

// 新規登録モーダルを閉じる
function closeAddModal() {
  document.getElementById('add-modal').classList.remove('show');
}

// 日報を登録
async function addReport() {
  const data = {
    name: document.getElementById('add-name').value,
    date: document.getElementById('add-date').value,
    tasks: [{
      category: document.getElementById('add-category').value,
      taskName: document.getElementById('add-task-name').value,
      workContent: document.getElementById('add-work-content').value,
      remarks: document.getElementById('add-remarks').value,
      workHours: parseFloat(document.getElementById('add-work-hours').value)
    }]
  };

  showLoading();

  try {
    const scriptUrl = getScriptUrl();
    const response = await fetch(scriptUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: `action=add&data=${encodeURIComponent(JSON.stringify(data))}`
    });

    const result = await response.json();

    if (result.status === 'success') {
      showToast('日報を登録しました', 'success');
      closeAddModal();

      // 検索条件が入力されていれば再検索
      const searchName = document.getElementById('search-name').value;
      if (searchName) {
        searchReports();
      }
    } else {
      throw new Error(result.message || '登録に失敗しました');
    }

  } catch (error) {
    console.error('Error:', error);
    showToast('登録に失敗しました: ' + error.message, 'error');
  } finally {
    hideLoading();
  }
}

// 編集モーダルを開く
function openEditModal(index) {
  const data = allReportsData[index];
  currentEditIndex = index;

  // フォームに値を設定
  document.getElementById('edit-row-index').value = index;
  document.getElementById('edit-name').value = data.name;
  document.getElementById('edit-date').value = formatDateToYYYYMMDD(new Date(data.date));
  document.getElementById('edit-category').value = data.category;
  document.getElementById('edit-task-name').value = data.taskName;
  document.getElementById('edit-work-content').value = data.workContent;
  document.getElementById('edit-remarks').value = data.remarks || '';
  document.getElementById('edit-work-hours').value = data.workHours;

  // モーダルを表示
  document.getElementById('edit-modal').classList.add('show');
}

// 編集モーダルを閉じる
function closeEditModal() {
  document.getElementById('edit-modal').classList.remove('show');
  currentEditIndex = null;
}

// 日報を更新
async function updateReport() {
  if (currentEditIndex === null) return;

  const originalData = allReportsData[currentEditIndex];
  const updatedData = {
    originalName: originalData.name,
    originalDate: formatDateToYYYYMMDD(new Date(originalData.date)),
    originalCategory: originalData.category,
    originalTaskName: originalData.taskName,
    originalWorkContent: originalData.workContent,
    originalRemarks: originalData.remarks || '',
    originalWorkHours: originalData.workHours,
    name: document.getElementById('edit-name').value,
    date: document.getElementById('edit-date').value,
    category: document.getElementById('edit-category').value,
    taskName: document.getElementById('edit-task-name').value,
    workContent: document.getElementById('edit-work-content').value,
    remarks: document.getElementById('edit-remarks').value,
    workHours: parseFloat(document.getElementById('edit-work-hours').value)
  };

  showLoading();

  try {
    const scriptUrl = getScriptUrl();
    const response = await fetch(scriptUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: `action=update&data=${encodeURIComponent(JSON.stringify(updatedData))}`
    });

    const result = await response.json();

    if (result.status === 'success') {
      showToast('日報を更新しました', 'success');
      closeEditModal();
      searchReports(); // 再検索
    } else {
      throw new Error(result.message || '更新に失敗しました');
    }

  } catch (error) {
    console.error('Error:', error);
    showToast('更新に失敗しました: ' + error.message, 'error');
  } finally {
    hideLoading();
  }
}

// 削除モーダルを開く
function openDeleteModal(index) {
  const data = allReportsData[index];
  currentDeleteIndex = index;

  // 削除する日報の情報を表示
  const displayDate = formatDateToJapanese(new Date(data.date));
  const deleteInfo = document.getElementById('delete-info');
  deleteInfo.innerHTML = `
    <strong>日付:</strong> ${displayDate}<br>
    <strong>分類:</strong> ${data.category}<br>
    <strong>タスク名:</strong> ${data.taskName}<br>
    <strong>作業内容:</strong> ${data.workContent}<br>
    <strong>時間:</strong> ${data.workHours}時間
  `;

  // モーダルを表示
  document.getElementById('delete-modal').classList.add('show');
}

// 削除モーダルを閉じる
function closeDeleteModal() {
  document.getElementById('delete-modal').classList.remove('show');
  currentDeleteIndex = null;
}

// 日報を削除
async function deleteReport() {
  if (currentDeleteIndex === null) return;

  const data = allReportsData[currentDeleteIndex];
  const deleteData = {
    name: data.name,
    date: formatDateToYYYYMMDD(new Date(data.date)),
    category: data.category,
    taskName: data.taskName,
    workContent: data.workContent,
    remarks: data.remarks || '',
    workHours: data.workHours
  };

  showLoading();

  try {
    const scriptUrl = getScriptUrl();
    const response = await fetch(scriptUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: `action=delete&data=${encodeURIComponent(JSON.stringify(deleteData))}`
    });

    const result = await response.json();

    if (result.status === 'success') {
      showToast('日報を削除しました', 'success');
      closeDeleteModal();
      searchReports(); // 再検索
    } else {
      throw new Error(result.message || '削除に失敗しました');
    }

  } catch (error) {
    console.error('Error:', error);
    showToast('削除に失敗しました: ' + error.message, 'error');
  } finally {
    hideLoading();
  }
}

// 検索条件をリセット
function resetSearch() {
  document.getElementById('search-name').value = '';
  document.querySelector('input[name="search-type"][value="period"]').checked = true;
  toggleSearchType('period');
  setDefaultDates();
  document.getElementById('results-container').innerHTML = '<p class="no-data">検索条件を入力して「検索」ボタンをクリックしてください</p>';
  document.getElementById('result-count').textContent = '0件';
  allReportsData = [];
  document.getElementById('btn-export-excel').disabled = true;
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

// 日付フォーマット（YYYY-MM-DD）
function formatDateToYYYYMMDD(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

// 日付フォーマット（YYYY年MM月DD日）
function formatDateToJapanese(date) {
  const year = date.getFullYear();
  const month = date.getMonth() + 1;
  const day = date.getDate();
  return `${year}年${month}月${day}日`;
}

// Excel出力
async function exportToExcel() {
  if (allReportsData.length === 0) {
    showToast('出力するデータがありません', 'warning');
    return;
  }

  showLoading();

  try {
    // ワークブックを作成
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('日報一覧');

    // 検索条件を取得
    const searchName = document.getElementById('search-name').value;
    const searchType = document.querySelector('input[name="search-type"]:checked').value;

    let startDate, endDate, conditionText;

    if (searchType === 'period') {
      startDate = document.getElementById('start-date').value;
      endDate = document.getElementById('end-date').value;
      conditionText = `検索条件： ${searchName}　期間： ${startDate} 〜 ${endDate}`;
    } else {
      const singleDate = document.getElementById('single-date').value;
      startDate = singleDate;
      endDate = singleDate;
      conditionText = `検索条件： ${searchName}　日付： ${singleDate}`;
    }

    // タイトル行
    worksheet.mergeCells('A1:F1');
    const titleCell = worksheet.getCell('A1');
    titleCell.value = '日報一覧';
    titleCell.font = { size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
    titleCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF667eea' }
    };
    titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getRow(1).height = 30;

    // 検索条件行
    worksheet.mergeCells('A2:F2');
    const conditionCell = worksheet.getCell('A2');
    conditionCell.value = conditionText;
    conditionCell.font = { size: 11 };
    conditionCell.alignment = { vertical: 'middle', horizontal: 'left' };
    worksheet.getRow(2).height = 20;

    // 空行
    worksheet.getRow(3).height = 10;

    // ヘッダー行
    const headerRow = worksheet.getRow(4);
    const headers = ['日付', '分類', 'タスク名', '作業内容', '備考', '時間'];
    headerRow.values = headers;
    headerRow.height = 25;

    // ヘッダーのスタイル（A4:F4の範囲のみ）
    for (let col = 1; col <= 6; col++) {
      const cell = headerRow.getCell(col);
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF764ba2' }
      };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    }

    // 列幅の設定
    worksheet.getColumn(1).width = 12;  // 日付
    worksheet.getColumn(2).width = 15;  // 分類
    worksheet.getColumn(3).width = 20;  // タスク名
    worksheet.getColumn(4).width = 40;  // 作業内容
    worksheet.getColumn(5).width = 30;  // 備考
    worksheet.getColumn(6).width = 10;  // 時間

    // データ行
    allReportsData.forEach((item, index) => {
      const rowIndex = index + 5;
      const row = worksheet.getRow(rowIndex);

      const displayDate = formatDateToJapanese(new Date(item.date));

      row.values = [
        displayDate,
        item.category,
        item.taskName,
        item.workContent,
        item.remarks || '',
        item.workHours
      ];

      // データ行のスタイル
      row.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
      row.height = 20;

      // 交互に背景色を設定
      if (index % 2 === 0) {
        row.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFF8F9FA' }
        };
      }

      // 枠線を設定（A列〜F列のみ）
      for (let col = 1; col <= 6; col++) {
        const cell = row.getCell(col);
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFE0E0E0' } },
          left: { style: 'thin', color: { argb: 'FFE0E0E0' } },
          bottom: { style: 'thin', color: { argb: 'FFE0E0E0' } },
          right: { style: 'thin', color: { argb: 'FFE0E0E0' } }
        };
      }
    });

    // ヘッダー行にも枠線を設定（A4:F4の範囲のみ）
    for (let col = 1; col <= 6; col++) {
      const cell = headerRow.getCell(col);
      cell.border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } }
      };
    }

    // 合計行を追加
    const totalRowIndex = allReportsData.length + 5;
    const totalRow = worksheet.getRow(totalRowIndex);
    const totalHours = allReportsData.reduce((sum, item) => sum + parseFloat(item.workHours), 0);

    totalRow.values = ['', '', '', '', '合計', totalHours.toFixed(1)];
    totalRow.height = 25;

    // 合計行のスタイルと枠線を設定（A列〜F列のみ）
    for (let col = 1; col <= 6; col++) {
      const cell = totalRow.getCell(col);
      cell.font = { bold: true };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFEAA7' }
      };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } }
      };
    }

    // ファイル名を生成
    let fileName;
    if (startDate === endDate) {
      fileName = `日報一覧_${searchName}_${startDate}.xlsx`;
    } else {
      fileName = `日報一覧_${searchName}_${startDate}_${endDate}.xlsx`;
    }

    // Excelファイルを生成してダウンロード
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });

    // ダウンロード
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);

    showToast('Excelファイルを出力しました', 'success');

  } catch (error) {
    console.error('Excel出力エラー:', error);
    showToast('Excel出力に失敗しました', 'error');
  } finally {
    hideLoading();
  }
}
