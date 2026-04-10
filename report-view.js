// Google Apps ScriptのウェブアプリURL
const GOOGLE_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyf8fJx5znNDEFfXxpaZE75HiWs8QSM3BKuZuPu_iJ3UWNwtjqrMpSzxAHg2sAXBLfSyA/exec';

// グローバル変数（エクスポート用にデータを保持）
let currentResults = null;
let currentDetailData = null;

// ページ読み込み時の初期化
document.addEventListener('DOMContentLoaded', function() {
    loadFilterOptions();
    setupEventListeners();
    setInitialValues();
});

// config.jsから選択肢を読み込み
function loadFilterOptions() {
    // 選択肢の読み込みは不要
}

// イベントリスナーの設定
function setupEventListeners() {
    document.getElementById('periodType').addEventListener('change', togglePeriodGroup);
    document.getElementById('reportSearchForm').addEventListener('submit', handleSearch);
    document.getElementById('reportSearchForm').addEventListener('reset', handleReset);
}

// 初期値の設定
function setInitialValues() {
    const today = new Date();

    // 期間選択のデフォルト値を設定
    const monthPicker = document.getElementById('monthPicker');
    monthPicker.value = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}`;

    const dayPicker = document.getElementById('dayPicker');
    dayPicker.value = today.toISOString().split('T')[0];

    const weekPicker = document.getElementById('weekPicker');
    const weekString = getWeekString(today);
    weekPicker.value = weekString;

    // 期間単位は「全期間」（空）をデフォルトにする
    document.getElementById('periodType').value = '';
}

// 週の文字列を取得（YYYY-Www形式）
function getWeekString(date) {
    const year = date.getFullYear();
    const onejan = new Date(year, 0, 1);
    const weekNumber = Math.ceil((((date - onejan) / 86400000) + onejan.getDay() + 1) / 7);
    return `${year}-W${String(weekNumber).padStart(2, '0')}`;
}

// 週番号から月と月内の週番号を取得（例: "2026-W15" → "4月2週目"）
function getWeekInfo(weekString) {
    const [year, week] = weekString.split('-W');
    const yearNum = parseInt(year);
    const weekNum = parseInt(week);

    // その年の1月1日
    const jan1 = new Date(yearNum, 0, 1);
    const jan1Day = jan1.getDay() || 7; // 日曜日を7に変換

    // その年の第1週の月曜日を計算
    const firstMonday = new Date(jan1);
    if (jan1Day <= 4) {
        // 1月1日が木曜日以前なら、その週が第1週
        firstMonday.setDate(jan1.getDate() - jan1Day + 1);
    } else {
        // 1月1日が金曜日以降なら、次の月曜日が第1週の開始
        firstMonday.setDate(jan1.getDate() + (8 - jan1Day));
    }

    // 対象週の月曜日
    const targetMonday = new Date(firstMonday);
    targetMonday.setDate(firstMonday.getDate() + (weekNum - 1) * 7);

    const month = targetMonday.getMonth() + 1; // 1-12

    // その月の1日を含む週の月曜日を取得
    const firstDayOfMonth = new Date(targetMonday.getFullYear(), targetMonday.getMonth(), 1);
    const firstDayDay = firstDayOfMonth.getDay() || 7; // 日曜日を7に変換

    // その月の1日を含む週の月曜日
    const firstWeekMonday = new Date(firstDayOfMonth);
    firstWeekMonday.setDate(1 - firstDayDay + 1);

    // 月内の週番号を計算（1日を含む週が第1週）
    const daysDiff = Math.floor((targetMonday - firstWeekMonday) / (1000 * 60 * 60 * 24));
    const weekOfMonth = Math.floor(daysDiff / 7) + 1;

    return { month, weekOfMonth };
}

// 期間グループの表示/非表示
function togglePeriodGroup() {
    const periodType = document.getElementById('periodType').value;
    document.getElementById('monthGroup').style.display = 'none';
    document.getElementById('weekGroup').style.display = 'none';
    document.getElementById('dayGroup').style.display = 'none';

    if (periodType === 'month') {
        document.getElementById('monthGroup').style.display = 'block';
    } else if (periodType === 'week') {
        document.getElementById('weekGroup').style.display = 'block';
    } else if (periodType === 'day') {
        document.getElementById('dayGroup').style.display = 'block';
    }
}


// ========================================
// データ取得・検索・集計処理
// ========================================

// 検索・集計処理
async function handleSearch(event) {
    event.preventDefault();

    const searchParams = getSearchParams();
    console.log('検索条件:', searchParams);

    // チェックボックスのバリデーション
    if (!searchParams.groupBy || searchParams.groupBy.length === 0) {
        alert('集計単位を1つ以上選択してください');
        return;
    }

    const resultsSection = document.getElementById('resultsSection');
    const resultsContainer = document.getElementById('resultsContainer');
    resultsSection.style.display = 'block';
    resultsContainer.innerHTML = '<div class="loading">データを取得中...</div>';

    try {
        // Google Sheetsからデータを取得
        const allData = await fetchDataFromGoogleSheets();

        // データをフィルタリング
        const filteredData = filterData(allData, searchParams);

        // 選択された集計方法の組み合わせに応じてデータを集計
        const aggregatedData = aggregateDataByCombination(filteredData, searchParams.groupBy);

        // 分類×タスクごとに備考をグループ化してサマリーを生成
        let remarksSummaryByGroup = null;
        remarksSummaryByGroup = await generateRemarksSummaryByGroup(filteredData);

        // 結果を表示
        displayResults(aggregatedData, filteredData, searchParams, remarksSummaryByGroup);

        // エクスポート用にデータを保存
        currentResults = aggregatedData;
        currentDetailData = filteredData;

        // 自動的にExcelエクスポートを実行
        if (filteredData.length > 0) {
            await autoExport(searchParams, aggregatedData, remarksSummaryByGroup);
        }

    } catch (error) {
        console.error('検索エラー:', error);
        resultsContainer.innerHTML = `
            <div style="padding: 20px; text-align: center; color: #e74c3c;">
                <p>❌ エラーが発生しました</p>
                <p style="font-size: 14px; margin-top: 8px;">${error.message}</p>
            </div>
        `;
    }
}

// Google Sheetsからデータを取得
async function fetchDataFromGoogleSheets() {
    try {
        console.log('Google Apps Script URLに接続中:', GOOGLE_SCRIPT_URL);

        const response = await fetch(GOOGLE_SCRIPT_URL, {
            method: 'GET',
            mode: 'cors'
        });

        console.log('レスポンスステータス:', response.status);

        if (!response.ok) {
            throw new Error(`HTTPエラー: ${response.status} ${response.statusText}`);
        }

        const result = await response.json();
        console.log('取得したデータ件数:', result.length);

        if (result.status === 'error') {
            throw new Error(result.message);
        }

        return result;
    } catch (error) {
        console.error('データ取得エラー:', error);
        throw new Error(`データ取得に失敗しました: ${error.message}`);
    }
}

// Gemini APIを使用して備考サマリーを生成
async function generateRemarksSummary(remarks) {
    try {
        console.log('備考サマリーを生成中...', remarks.length + '件');

        const url = `${GOOGLE_SCRIPT_URL}?action=summarize&remarks=${encodeURIComponent(JSON.stringify(remarks))}`;

        const response = await fetch(url, {
            method: 'GET',
            mode: 'cors'
        });

        if (!response.ok) {
            throw new Error(`HTTPエラー: ${response.status}`);
        }

        const result = await response.json();

        if (result.success) {
            console.log('✅ 備考サマリー生成成功');
            return result.summary;
        } else {
            console.warn('⚠️ 備考サマリー生成失敗:', result.message);
            return `備考サマリーの生成に失敗しました: ${result.message}`;
        }

    } catch (error) {
        console.error('備考サマリー生成エラー:', error);
        return `備考サマリーの生成エラー: ${error.message}`;
    }
}

// 分類ごとに備考をグループ化してサマリーを生成
async function generateRemarksSummaryByGroup(data) {
    try {
        console.log('分類ごとに備考サマリーを生成中...');

        // 分類ごとに備考をグループ化
        const groupedRemarks = {};

        data.forEach(row => {
            const category = row.category || '未分類';
            const remarks = row.remarks ? row.remarks.trim() : '';

            if (remarks !== '') {
                if (!groupedRemarks[category]) {
                    groupedRemarks[category] = {
                        category: category,
                        remarks: []
                    };
                }
                groupedRemarks[category].remarks.push(remarks);
            }
        });

        // グループごとにサマリーを生成
        const summaryResults = [];

        for (const category in groupedRemarks) {
            const group = groupedRemarks[category];

            // 重複を除去
            const uniqueRemarks = [...new Set(group.remarks)];

            console.log(`[${group.category}] 備考: ${uniqueRemarks.length}件`);

            // サマリーを生成
            const summary = await generateRemarksSummary(uniqueRemarks);

            summaryResults.push({
                category: group.category,
                remarksCount: uniqueRemarks.length,
                summary: summary
            });

            // レート制限対策: 1秒待機
            await new Promise(resolve => setTimeout(resolve, 1000));
        }

        console.log(`✅ ${summaryResults.length}グループのサマリー生成完了`);
        return summaryResults;

    } catch (error) {
        console.error('グループ別サマリー生成エラー:', error);
        return null;
    }
}

// データをフィルタリング
function filterData(data, params) {
    console.log('=== フィルタリング開始 ===');
    console.log('全データ件数:', data.length);
    console.log('検索条件:', params);

    const filteredData = data.filter(row => {
        // 期間フィルター
        if (params.period.type && params.period.value) {
            const rowDate = new Date(row.date);

            if (params.period.type === 'day') {
                let normalizedRowDate = row.date;

                // 日付が文字列でない場合（Dateオブジェクトの場合）
                if (normalizedRowDate instanceof Date) {
                    const year = normalizedRowDate.getFullYear();
                    const month = String(normalizedRowDate.getMonth() + 1).padStart(2, '0');
                    const day = String(normalizedRowDate.getDate()).padStart(2, '0');
                    normalizedRowDate = `${year}-${month}-${day}`;
                } else if (typeof normalizedRowDate === 'string') {
                    // ISO形式（YYYY-MM-DDTHH:mm:ss.sssZ）の場合はDateオブジェクトに変換してローカル日付を取得
                    if (normalizedRowDate.includes('T') && (normalizedRowDate.includes('Z') || normalizedRowDate.includes('+'))) {
                        const dateObj = new Date(normalizedRowDate);
                        const year = dateObj.getFullYear();
                        const month = String(dateObj.getMonth() + 1).padStart(2, '0');
                        const day = String(dateObj.getDate()).padStart(2, '0');
                        normalizedRowDate = `${year}-${month}-${day}`;
                    }
                    // 単純なISO形式（YYYY-MM-DDTHH:mm:ss）の場合
                    else if (normalizedRowDate.includes('T')) {
                        normalizedRowDate = normalizedRowDate.split('T')[0];
                    }
                    // スラッシュ区切り（YYYY/MM/DD）の場合
                    else if (normalizedRowDate.includes('/')) {
                        const parts = normalizedRowDate.split('/');
                        if (parts.length === 3) {
                            normalizedRowDate = `${parts[0]}-${parts[1].padStart(2, '0')}-${parts[2].padStart(2, '0')}`;
                        }
                    }
                    normalizedRowDate = normalizedRowDate.trim();
                }

                const normalizedSearchDate = params.period.value.trim();

                console.log(`日付比較: 元="${row.date}" (型:${typeof row.date}) → 正規化="${normalizedRowDate}" vs 検索="${normalizedSearchDate}" → 一致:${normalizedRowDate === normalizedSearchDate}`);

                if (normalizedRowDate !== normalizedSearchDate) return false;
            } else if (params.period.type === 'month') {
                const [year, month] = params.period.value.split('-');
                const rowYear = rowDate.getFullYear();
                const rowMonth = rowDate.getMonth() + 1;
                if (rowYear !== parseInt(year) || rowMonth !== parseInt(month)) return false;
            } else if (params.period.type === 'week') {
                const weekString = getWeekString(rowDate);
                if (weekString !== params.period.value) return false;
            }
        }

        return true;
    });

    console.log('フィルタリング後の件数:', filteredData.length);
    console.log('=== フィルタリング終了 ===');

    return filteredData;
}

// 選択された組み合わせに応じてデータを集計（ピボットテーブル形式）
function aggregateDataByCombination(data, groupByArray) {
    // 集計キーの順序を決定（name → category → task）
    const groupByOrder = [];
    if (groupByArray.includes('name')) groupByOrder.push('name');
    if (groupByArray.includes('category')) groupByOrder.push('category');
    if (groupByArray.includes('task')) groupByOrder.push('task');

    // 階層構造で集計
    const hierarchy = {};

    data.forEach(row => {
        const keys = groupByOrder.map(field => {
            if (field === 'name') return row.name;
            if (field === 'category') return row.category;
            if (field === 'task') return row.taskName;
            return '';
        });

        const hours = parseFloat(row.workHours || 0);

        // 階層構造を構築
        let current = hierarchy;
        keys.forEach((key, index) => {
            if (!current[key]) {
                current[key] = { _hours: 0, _children: {} };
            }
            current[key]._hours += hours;
            if (index < keys.length - 1) {
                current = current[key]._children;
            }
        });
    });

    // ピボットテーブル形式のデータを作成
    const pivotData = [];
    let grandTotal = 0;

    function processLevel(obj, level, parentKeys) {
        const keys = Object.keys(obj).sort();
        keys.forEach(key => {
            const item = obj[key];
            const currentKeys = [...parentKeys, key];

            // 最下層の場合は、そのまま行を追加
            if (Object.keys(item._children).length === 0) {
                const rowData = { hours: item._hours, isSubtotal: false, isGrandTotal: false };
                groupByOrder.forEach((field, index) => {
                    rowData[field] = currentKeys[index] || '';
                });
                pivotData.push(rowData);
            } else {
                // 子要素を処理
                processLevel(item._children, level + 1, currentKeys);

                // 小計行を追加
                const subtotalData = { hours: item._hours, isSubtotal: true, isGrandTotal: false, subtotalLevel: level };
                groupByOrder.forEach((field, index) => {
                    if (index <= level) {
                        subtotalData[field] = currentKeys[index] || '';
                    } else {
                        subtotalData[field] = '';
                    }
                });
                pivotData.push(subtotalData);
            }

            grandTotal = Math.max(grandTotal, item._hours);
        });
    }

    // 総合計を計算
    let calculatedGrandTotal = 0;
    Object.keys(hierarchy).forEach(key => {
        calculatedGrandTotal += hierarchy[key]._hours;
    });

    processLevel(hierarchy, 0, []);

    // 総合計行を追加
    const grandTotalData = { hours: calculatedGrandTotal, isSubtotal: false, isGrandTotal: true };
    groupByOrder.forEach((field, index) => {
        if (index === 0) {
            grandTotalData[field] = '総合計';
        } else {
            grandTotalData[field] = '';
        }
    });
    pivotData.push(grandTotalData);

    return {
        type: 'combination',
        groupByOrder: groupByOrder,
        pivotData: pivotData
    };
}


// 結果を表示
function displayResults(aggregatedData, detailData, params, remarksSummary) {
    const resultsContainer = document.getElementById('resultsContainer');

    if (detailData.length === 0) {
        resultsContainer.innerHTML = `
            <div style="padding: 40px; text-align: center; color: #999;">
                <p style="font-size: 18px;">📭 該当するデータがありません</p>
                <p style="font-size: 14px; margin-top: 12px;">検索条件を変更してください</p>
            </div>
        `;
        return;
    }

    let html = '<div class="results-wrapper">';

    // 集計期間
    html += '<div class="summary-section">';
    html += '<h3 class="summary-title">📅 集計期間</h3>';
    html += '<div class="period-info">';
    if (params.period.type === 'day') {
        html += `<strong>対象日:</strong> ${params.period.value}`;
    } else if (params.period.type === 'week') {
        const weekInfo = getWeekInfo(params.period.value);
        html += `<strong>対象週:</strong> ${params.period.value}（${weekInfo.month}月${weekInfo.weekOfMonth}週目）`;
    } else if (params.period.type === 'month') {
        html += `<strong>対象月:</strong> ${params.period.value}`;
    } else {
        html += `<strong>全期間</strong>`;
    }
    html += '</div>';
    html += '</div>';

    // ピボットテーブルを表示
    html += displayPivotTable(aggregatedData, params);

    // 分類ごとの備考サマリーを表示
    if (remarksSummary && remarksSummary.length > 0) {
        html += '<div class="summary-section">';
        html += '<h3 class="summary-title">🤖 備考サマリー（AI生成）</h3>';

        remarksSummary.forEach(group => {
            html += '<div class="remarks-group">';
            html += `<h4 class="remarks-group-title">【${group.category}】 備考${group.remarksCount}件</h4>`;
            html += '<div class="remarks-summary">';
            html += `<pre class="ai-summary">${group.summary}</pre>`;
            html += '</div>';
            html += '</div>';
        });

        html += '</div>';
    }

    html += '</div>';

    resultsContainer.innerHTML = html;
}

// ピボットテーブルを表示
function displayPivotTable(data, params) {
    let html = '';

    const groupByOrder = data.groupByOrder;

    // タイトル作成
    const titleParts = groupByOrder.map(field => {
        if (field === 'name') return '名前';
        if (field === 'category') return '分類';
        if (field === 'task') return 'タスク';
        return '';
    });
    const title = titleParts.join(' × ') + ' 集計';

    html += '<div class="summary-section">';
    html += `<h3 class="summary-title">📊 ${title}</h3>`;
    html += '<table class="summary-table">';

    // ヘッダー行
    html += '<thead><tr>';
    groupByOrder.forEach(field => {
        if (field === 'name') html += '<th>名前</th>';
        if (field === 'category') html += '<th>分類</th>';
        if (field === 'task') html += '<th>タスク名</th>';
    });
    html += '<th>作業時間</th>';
    html += '</tr></thead>';

    // データ行
    html += '<tbody>';

    const currentValues = {};
    groupByOrder.forEach(field => {
        currentValues[field] = '';
    });

    data.pivotData.forEach(item => {
        html += '<tr';
        if (item.isSubtotal) {
            html += ' style="background: #fff9e6; font-weight: bold;"';
        } else if (item.isGrandTotal) {
            html += ' style="background: #f0f4ff; font-weight: bold;"';
        }
        html += '>';

        // 各列を出力
        groupByOrder.forEach((field, index) => {
            const value = item[field] || '';

            if (item.isSubtotal && index === item.subtotalLevel) {
                // その階層の項目名のみを表示
                const subtotalValue = item[groupByOrder[item.subtotalLevel]];
                html += `<td style="text-align: right; padding-right: 16px;">${subtotalValue} 合計</td>`;
            } else if (item.isGrandTotal && index === 0) {
                html += `<td style="text-align: right; padding-right: 16px;">総合計</td>`;
            } else if (item.isSubtotal || item.isGrandTotal) {
                html += '<td></td>';
            } else {
                // 同じ値が続く場合は空欄
                if (value === currentValues[field]) {
                    html += '<td></td>';
                } else {
                    html += `<td>${value}</td>`;
                    currentValues[field] = value;
                    // それ以降の階層をリセット
                    const resetFrom = index + 1;
                    for (let i = resetFrom; i < groupByOrder.length; i++) {
                        currentValues[groupByOrder[i]] = '';
                    }
                }
            }
        });

        // 作業時間列
        html += `<td class="hours-cell">${item.hours.toFixed(1)}h</td>`;

        html += '</tr>';
    });

    html += '</tbody></table>';
    html += '</div>';

    return html;
}


// 検索条件の取得
function getSearchParams() {
    // チェックボックスから選択された集計方法を取得
    const checkboxes = document.querySelectorAll('input[name="groupBy"]:checked');
    const groupByArray = Array.from(checkboxes).map(cb => cb.value);

    const params = {
        groupBy: groupByArray,
        period: {}
    };

    const periodType = document.getElementById('periodType').value;
    params.period.type = periodType;

    if (periodType === 'month') {
        params.period.value = document.getElementById('monthPicker').value;
    } else if (periodType === 'week') {
        params.period.value = document.getElementById('weekPicker').value;
    } else if (periodType === 'day') {
        params.period.value = document.getElementById('dayPicker').value;
    }

    return params;
}

// ========================================
// Excelエクスポート処理（ExcelJS使用）
// ========================================

// 自動エクスポート（検索・集計ボタン押下時）
async function autoExport(searchParams, aggregatedData, remarksSummary) {
    if (!aggregatedData || !currentDetailData) {
        return;
    }

    try {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = '日報システム';
        workbook.created = new Date();

        const borderStyle = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        // ピボットテーブルをエクスポート
        exportPivotTable(workbook, aggregatedData, searchParams, borderStyle, remarksSummary);

        // 集計方法の日本語名を取得（複数の場合は結合）
        const groupByNames = searchParams.groupBy.map(type => {
            if (type === 'name') return '名前';
            if (type === 'category') return '分類';
            if (type === 'task') return 'タスク';
            return '';
        }).join('_');

        // ファイル名生成：日報集計_選択した集計方法_日付時分秒
        const now = new Date();
        const filename = `日報集計_${groupByNames}_${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}${String(now.getSeconds()).padStart(2, '0')}.xlsx`;

        // ダウンロード
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = filename;
        link.click();
        window.URL.revokeObjectURL(url);

        console.log('✅ Excelファイルをダウンロードしました:', filename);

    } catch (error) {
        console.error('エクスポートエラー:', error);
        alert('エクスポートに失敗しました: ' + error.message);
    }
}

// ピボットテーブルをExcelにエクスポート
function exportPivotTable(workbook, data, params, borderStyle, remarksSummary) {
    const groupByOrder = data.groupByOrder;

    // シート名を作成
    const sheetNameParts = groupByOrder.map(field => {
        if (field === 'name') return '名前';
        if (field === 'category') return '分類';
        if (field === 'task') return 'タスク';
        return '';
    });
    const sheetName = sheetNameParts.join('_') + '集計';

    const sheet = workbook.addWorksheet(sheetName);

    let currentRow = 1;

    // 集計期間
    const numCols = groupByOrder.length + 1; // 集計項目 + 作業時間
    const lastCol = String.fromCharCode(64 + numCols); // A=65
    sheet.mergeCells(`A${currentRow}:${lastCol}${currentRow}`);
    const periodCell = sheet.getCell(`A${currentRow}`);
    periodCell.value = '集計期間';
    periodCell.font = { size: 14, bold: true };
    periodCell.alignment = { vertical: 'middle', horizontal: 'center' };
    sheet.getRow(currentRow).height = 25;
    currentRow++;

    let periodText = '';
    if (params.period.type === 'day') {
        periodText = `対象日: ${params.period.value}`;
    } else if (params.period.type === 'week') {
        const weekInfo = getWeekInfo(params.period.value);
        periodText = `対象週: ${params.period.value}（${weekInfo.month}月${weekInfo.weekOfMonth}週目）`;
    } else if (params.period.type === 'month') {
        periodText = `対象月: ${params.period.value}`;
    } else {
        periodText = '全期間';
    }
    sheet.mergeCells(`A${currentRow}:${lastCol}${currentRow}`);
    const periodValueCell = sheet.getCell(`A${currentRow}`);
    periodValueCell.value = periodText;
    currentRow += 2;

    // タイトル
    sheet.mergeCells(`A${currentRow}:${lastCol}${currentRow}`);
    const titleCell = sheet.getCell(`A${currentRow}`);
    titleCell.value = sheetNameParts.join(' × ') + ' 集計';
    titleCell.font = { size: 12, bold: true };
    titleCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE0E0E0' }
    };
    currentRow++;

    // ヘッダー行
    const headerData = [];
    groupByOrder.forEach(field => {
        if (field === 'name') headerData.push('名前');
        if (field === 'category') headerData.push('分類');
        if (field === 'task') headerData.push('タスク名');
    });
    headerData.push('作業時間');

    const headerRow = sheet.addRow(headerData);
    headerRow.eachCell(cell => {
        cell.font = { bold: true };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFB0E0E6' }
        };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border = borderStyle;
    });

    // データ行
    const currentValues = {};
    groupByOrder.forEach(field => {
        currentValues[field] = '';
    });

    data.pivotData.forEach(item => {
        const rowData = [];

        // 各列のデータ
        groupByOrder.forEach((field, index) => {
            const value = item[field] || '';

            if (item.isSubtotal && index === item.subtotalLevel) {
                // その階層の項目名のみを表示
                const subtotalValue = item[groupByOrder[item.subtotalLevel]];
                rowData.push(subtotalValue + ' 合計');
            } else if (item.isGrandTotal && index === 0) {
                rowData.push('総合計');
            } else if (item.isSubtotal || item.isGrandTotal) {
                rowData.push('');
            } else {
                // 同じ値が続く場合は空欄
                if (value === currentValues[field]) {
                    rowData.push('');
                } else {
                    rowData.push(value);
                    currentValues[field] = value;
                    // それ以降の階層をリセット
                    const resetFrom = index + 1;
                    for (let i = resetFrom; i < groupByOrder.length; i++) {
                        currentValues[groupByOrder[i]] = '';
                    }
                }
            }
        });

        // 作業時間
        rowData.push(`${item.hours.toFixed(1)}h`);

        const row = sheet.addRow(rowData);

        row.eachCell((cell, colNumber) => {
            cell.border = borderStyle;
            if (item.isSubtotal) {
                cell.font = { bold: true };
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFFFF9E6' }
                };
                // 小計行の階層の列を右寄せ
                if (colNumber === item.subtotalLevel + 1) {
                    cell.alignment = { vertical: 'middle', horizontal: 'right' };
                }
            } else if (item.isGrandTotal) {
                cell.font = { bold: true };
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFF0F4FF' }
                };
                if (colNumber === 1) {
                    cell.alignment = { vertical: 'middle', horizontal: 'right' };
                }
            }
        });
    });

    // 列幅調整
    for (let i = 1; i <= numCols; i++) {
        sheet.getColumn(i).width = 20;
    }

    // 備考サマリーを追加（分類ごと）
    if (remarksSummary && remarksSummary.length > 0) {
        currentRow = sheet.lastRow.number + 2;

        sheet.mergeCells(`A${currentRow}:${lastCol}${currentRow}`);
        const summaryTitle = sheet.getCell(`A${currentRow}`);
        summaryTitle.value = '備考サマリー（AI生成）';
        summaryTitle.font = { size: 12, bold: true };
        summaryTitle.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE0E0E0' }
        };
        currentRow++;

        // 各グループのサマリーを出力
        remarksSummary.forEach(group => {
            // グループタイトル
            sheet.mergeCells(`A${currentRow}:${lastCol}${currentRow}`);
            const groupTitle = sheet.getCell(`A${currentRow}`);
            groupTitle.value = `【${group.category}】 備考${group.remarksCount}件`;
            groupTitle.font = { size: 11, bold: true };
            groupTitle.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFF0F4FF' }
            };
            groupTitle.alignment = { vertical: 'middle', horizontal: 'left' };
            groupTitle.border = borderStyle;
            currentRow++;

            // サマリー内容を行に分割して出力
            const summaryLines = group.summary.split('\n');
            summaryLines.forEach(line => {
                sheet.mergeCells(`A${currentRow}:${lastCol}${currentRow}`);
                const summaryCell = sheet.getCell(`A${currentRow}`);
                summaryCell.value = line;
                summaryCell.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
                summaryCell.border = borderStyle;
                currentRow++;
            });

            // グループ間に空行
            currentRow++;
        });

        // サマリー列の幅を広く
        sheet.getColumn(1).width = 80;
    }
}

// リセット処理
function handleReset() {
    document.getElementById('monthGroup').style.display = 'none';
    document.getElementById('weekGroup').style.display = 'none';
    document.getElementById('dayGroup').style.display = 'none';
    document.getElementById('resultsSection').style.display = 'none';

    currentResults = null;
    currentDetailData = null;

    setTimeout(setInitialValues, 0);
}

