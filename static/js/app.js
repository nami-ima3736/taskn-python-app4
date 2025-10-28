// 在留資格管理システム - JavaScript

// グローバル変数
let allData = [];
let filteredData = [];
let currentFilter = 'all';
let currentSort = { column: null, ascending: true };
let hasUnsavedChanges = false;
let lastChangeTimestamp = null;
let lastExportedAt = null;

// 期生のフォーマット関数（数値のみの場合は「期」を追加）
function formatKisei(kisei) {
    if (!kisei || kisei === '-') {
        return '-';
    }
    const kiseiStr = String(kisei).trim();
    // 数値のみの場合は「期」を追加
    if (/^\d+$/.test(kiseiStr)) {
        return kiseiStr + '期';
    }
    return kiseiStr;
}

// ページ読み込み時の初期化
document.addEventListener('DOMContentLoaded', function() {
    // ファイル選択イベント
    document.getElementById('fileInput').addEventListener('change', handleFileUpload);
    
    // フォーム送信イベント
    document.getElementById('editForm').addEventListener('submit', handleFormSubmit);
    
    // 初期データ読み込み
    loadData();
    loadSummary();

    updateUnsavedIndicator();
});

window.addEventListener('beforeunload', (event) => {
    if (hasUnsavedChanges) {
        event.preventDefault();
        event.returnValue = '未保存の変更があります。アプリを終了しますか？';
    }
});

// データを読み込む
async function loadData() {
    try {
        updateStatus('データを読み込み中...');
        const response = await fetch('/api/data');
        
        if (!response.ok) {
            throw new Error('データの読み込みに失敗しました');
        }
        
        const result = await response.json();
        allData = result.data;
        filteredData = allData;
        
        updateTable();
        updateStatus(`データを読み込みました (${result.total}件)`);
    } catch (error) {
        console.error('Error loading data:', error);
        updateStatus('エラー: ' + error.message);
    }
}

// サマリー情報を読み込む
async function loadSummary() {
    try {
        const response = await fetch('/api/summary');
        
        if (!response.ok) {
            return;
        }
        
        const summary = await response.json();
        
        // ファイル情報を更新
        document.getElementById('filename').textContent = `ファイル: ${summary.filename}`;
        document.getElementById('totalCount').textContent = `総件数: ${summary.total}件`;
        updateLastExportInfo(summary);
        
        // 期限状況を更新
        document.getElementById('deadlinePassedCount').textContent = `${summary.deadline_passed}件`;
        document.getElementById('expiredCount').textContent = `${summary.expired}件`;
        document.getElementById('days30Count').textContent = `${summary.days_30_count}件`;
        document.getElementById('days60Count').textContent = `${summary.days_60_count}件`;
        document.getElementById('days90Count').textContent = `${summary.days_90_count}件`;
        document.getElementById('skill1LimitCount').textContent = `${summary.skill1_limit_count}件`;
    } catch (error) {
        console.error('Error loading summary:', error);
    }
}

// テーブルを更新
function updateTable() {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';
    
    if (filteredData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="16" class="no-data">表示するデータがありません</td></tr>';
        return;
    }
    
    filteredData.forEach(row => {
        const tr = document.createElement('tr');
        
        // 行の背景色クラスを設定（満了日数に基づく）
        const rowClass = `row-${row['状態']}`;
        tr.className = rowClass;
        
        // 各セルを作成
        const cells = [
            { value: row['担当者コード'], isDeadline: false },
            { value: row['氏名１'], isDeadline: false },
            { value: row['氏名２'], isDeadline: false },
            { value: row['在留資格'], isDeadline: false },
            { value: row['国籍'], isDeadline: false },
            { value: row['在留カード番号'] || '-', isDeadline: false },
            { value: row['生年月日'] || '-', isDeadline: false },
            { value: formatKisei(row['期生']), isDeadline: false },
            { value: row['許可年月日'], isDeadline: false },
            { value: row['満了年月日'], isDeadline: false },
            { value: row['満了日数'], isDeadline: false },
            { value: row['期限日1'], isDeadline: true, status: row['期限日1_状態'] },
            { value: row['期限日2'], isDeadline: true, status: row['期限日2_状態'] },
            { value: row['期限日3'], isDeadline: true, status: row['期限日3_状態'] },
            { value: getStatusText(row['状態']), isDeadline: false }
        ];
        
        cells.forEach(cell => {
            const td = document.createElement('td');
            
            // 期限日セルの場合、超過していれば背景色を変更
            if (cell.isDeadline && cell.status && cell.status.status === 'overdue') {
                td.className = 'cell-deadline-overdue';
                td.textContent = `🔴⚠️ ${cell.value} (超過${cell.status.days}日)`;
            } else {
                td.textContent = cell.value;
            }
            
            tr.appendChild(td);
        });
        
        // 操作ボタンを追加
        const actionTd = document.createElement('td');
        actionTd.className = 'action-buttons';
        
        const editBtn = document.createElement('button');
        editBtn.textContent = '編集';
        editBtn.className = 'btn-small btn-edit';
        editBtn.onclick = () => showEditModal(row['index']);
        
        const deleteBtn = document.createElement('button');
        deleteBtn.textContent = '削除';
        deleteBtn.className = 'btn-small btn-delete';
        deleteBtn.onclick = () => deleteRow(row['index']);
        
        actionTd.appendChild(editBtn);
        actionTd.appendChild(deleteBtn);
        tr.appendChild(actionTd);
        
        tbody.appendChild(tr);
    });
}

// 状態テキストを取得
function getStatusText(status) {
    const statusMap = {
        'expired': '期限切れ',
        'urgent': '緊急',
        'warning': '警告',
        'caution': '注意',
        'safe': '安全',
        'unknown': '-'
    };
    return statusMap[status] || '-';
}

// データをフィルタリング
function filterData(filterType, event) {
    currentFilter = filterType;
    
    // フィルターボタンのアクティブ状態を更新
    document.querySelectorAll('.btn-filter').forEach(btn => {
        btn.classList.remove('active');
    });
    if (event && event.target) {
        event.target.classList.add('active');
    }
    
    if (filterType === 'all') {
        filteredData = allData;
        updateStatus('フィルター: すべて表示');
    } else if (filterType === 'deadline_passed') {
        // 期限日が超過しているデータのみ
        filteredData = allData.filter(row => {
            return (row['期限日1_状態'] && row['期限日1_状態'].status === 'overdue') ||
                   (row['期限日2_状態'] && row['期限日2_状態'].status === 'overdue') ||
                   (row['期限日3_状態'] && row['期限日3_状態'].status === 'overdue');
        });
        updateStatus(`フィルター: 期限日超過のみ (${filteredData.length}件)`);
    } else if (filterType === 'expired') {
        // 満了期限切れ（満了年月日が過ぎている）
        filteredData = allData.filter(row => {
            const days = row['満了年月日までの日数'];
            return days !== null && days < 0;
        });
        updateStatus(`フィルター: 満了期限切れのみ (${filteredData.length}件)`);
    } else if (filterType === '30days') {
        // レベル1（最も余裕が少ない分類）
        filteredData = allData.filter(row => {
            return row['期限レベル分類'] === 'level1';
        });
        updateStatus(`フィルター: 期限レベル1 (${filteredData.length}件)`);
    } else if (filterType === '60days') {
        // レベル2（中程度の余裕）
        filteredData = allData.filter(row => {
            return row['期限レベル分類'] === 'level2';
        });
        updateStatus(`フィルター: 期限レベル2 (${filteredData.length}件)`);
    } else if (filterType === '90days') {
        // レベル3（最も余裕がある分類）
        filteredData = allData.filter(row => {
            return row['期限レベル分類'] === 'level3';
        });
        updateStatus(`フィルター: 期限レベル3 (${filteredData.length}件)`);
    } else if (filterType === 'skill1_limit') {
        // 特定技能1号期限超過（満了日数 + 184 > 1826）
        filteredData = allData.filter(row => {
            const days = row['満了日数_値'];
            return days !== null && (days + 184) > 1826;
        });
        updateStatus(`フィルター: 特定技能1号期限超過 (${filteredData.length}件)`);
    }
    
    updateTable();
}

// データを検索
function searchData() {
    const searchText = document.getElementById('searchInput').value.toLowerCase().trim();
    
    if (!searchText) {
        filteredData = allData;
        document.getElementById('searchResult').textContent = '';
        updateTable();
        updateStatus('検索クリア');
        return;
    }
    
    filteredData = allData.filter(row => {
        return row['担当者コード'].toLowerCase().includes(searchText) ||
               row['氏名１'].toLowerCase().includes(searchText) ||
               row['氏名２'].toLowerCase().includes(searchText) ||
               row['在留資格'].toLowerCase().includes(searchText) ||
               row['国籍'].toLowerCase().includes(searchText);
    });
    
    document.getElementById('searchResult').textContent = `(${filteredData.length}件)`;
    updateTable();
    updateStatus(`検索結果: ${filteredData.length}件`);
}

// テーブルをソート
function sortTable(column) {
    if (currentSort.column === column) {
        currentSort.ascending = !currentSort.ascending;
    } else {
        currentSort.column = column;
        currentSort.ascending = true;
    }
    
    filteredData.sort((a, b) => {
        let aVal = a[column];
        let bVal = b[column];
        
        // 数値の場合
        if (column === '満了日数_値') {
            aVal = aVal === null ? Infinity : aVal;
            bVal = bVal === null ? Infinity : bVal;
        }
        
        if (aVal < bVal) return currentSort.ascending ? -1 : 1;
        if (aVal > bVal) return currentSort.ascending ? 1 : -1;
        return 0;
    });
    
    updateTable();
    updateStatus(`ソート: ${column} (${currentSort.ascending ? '昇順' : '降順'})`);
}

// データを更新
async function refreshData() {
    await loadData();
    await loadSummary();
}

// ファイルをアップロード
async function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    const formData = new FormData();
    formData.append('file', file);
    
    try {
        updateStatus('ファイルをアップロード中...');
        const response = await fetch('/api/upload', {
            method: 'POST',
            body: formData
        });
        
        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.error || 'アップロードに失敗しました');
        }
        
        const result = await response.json();
        updateStatus(`ファイルを読み込みました: ${result.filename}`);
        
        // データを再読み込み
        await loadData();
        await loadSummary();
        resetUnsavedChanges();
    } catch (error) {
        console.error('Error uploading file:', error);
        updateStatus('エラー: ' + error.message);
        alert('ファイルのアップロードに失敗しました:\n' + error.message);
    }
}

// アラートリストをエクスポート
async function exportAlertList() {
    try {
        updateStatus('アラートリストを出力中...');
        window.location.href = '/api/export/alert';
        updateStatus('アラートリストを出力しました');
    } catch (error) {
        console.error('Error exporting alert list:', error);
        updateStatus('エラー: アラートリストの出力に失敗しました');
        alert('アラートリストの出力に失敗しました');
    }
}

// 処理済みデータをエクスポート
async function exportProcessedData() {
    try {
        updateStatus('処理済みデータを保存中...');
        const response = await fetch('/api/export/processed');

        if (!response.ok) {
            const error = await response.json().catch(() => ({ error: '処理済みデータの保存に失敗しました' }));
            throw new Error(error.error || '処理済みデータの保存に失敗しました');
        }

        const disposition = response.headers.get('Content-Disposition') || '';
        const filename = extractFilenameFromDisposition(disposition) || '在留資格管理_processed.xlsx';
        const blob = await response.blob();

        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        window.URL.revokeObjectURL(url);

        updateStatus('処理済みデータを保存しました');

        resetUnsavedChanges();
        await loadSummary();
    } catch (error) {
        console.error('Error exporting processed data:', error);
        updateStatus('エラー: 処理済みデータの保存に失敗しました');
        alert('処理済みデータの保存に失敗しました:\n' + error.message);
    }
}

// ステータスバーを更新
function updateStatus(message) {
    document.getElementById('statusBar').textContent = message;
}

// データ追加モーダルを表示
function showAddModal() {
    document.getElementById('modalTitle').textContent = 'データ追加';
    document.getElementById('editIndex').value = '';
    
    // フォームをリセット
    document.getElementById('editForm').reset();
    
    // すべてのフィールドを明示的にクリア
    document.getElementById('edit担当者コード').value = '';
    document.getElementById('edit氏名１').value = '';
    document.getElementById('edit氏名２').value = '';
    document.getElementById('edit在留資格').value = '';
    document.getElementById('edit国籍').value = '';
    document.getElementById('edit在留カード番号').value = '';
    document.getElementById('edit生年月日').value = '';
    document.getElementById('edit期生').value = '';
    document.getElementById('edit許可年月日').value = '';
    document.getElementById('edit満了年月日').value = '';
    document.getElementById('edit既満了日数').value = '';  // 空にする（特定技能1号以外は不要）
    document.getElementById('edit設定期限1').value = '90';
    document.getElementById('edit設定期限2').value = '60';
    document.getElementById('edit設定期限3').value = '30';
    
    document.getElementById('editModal').style.display = 'block';
}

// データ編集モーダルを表示
async function showEditModal(index) {
    const row = allData.find(r => r.index === index);
    if (!row) {
        alert('データが見つかりません');
        return;
    }
    
    document.getElementById('modalTitle').textContent = 'データ編集';
    document.getElementById('editIndex').value = index;
    
    // フォームに値を設定
    document.getElementById('edit担当者コード').value = row['担当者コード'] || '';
    document.getElementById('edit氏名１').value = row['氏名１'] || '';
    document.getElementById('edit氏名２').value = row['氏名２'] || '';
    document.getElementById('edit在留資格').value = row['在留資格'] || '';
    document.getElementById('edit国籍').value = row['国籍'] || '';
    document.getElementById('edit在留カード番号').value = row['在留カード番号'] || '';
    document.getElementById('edit期生').value = row['期生'] ? String(row['期生']).replace(/[^0-9]/g, '') : '';
    
    // 日付をYYYY-MM-DD形式に変換
    const kyokaDate = row['許可年月日'];
    if (kyokaDate && kyokaDate !== '-') {
        document.getElementById('edit許可年月日').value = convertToDateInput(kyokaDate);
    }
    
    const manryoDate = row['満了年月日'];
    if (manryoDate && manryoDate !== '-') {
        document.getElementById('edit満了年月日').value = convertToDateInput(manryoDate);
    }
    
    const birthDate = row['生年月日'];
    if (birthDate && birthDate !== '-') {
        document.getElementById('edit生年月日').value = convertToDateInput(birthDate);
    }
    
    document.getElementById('editModal').style.display = 'block';
}

// 日付をYYYY-MM-DD形式に変換
function convertToDateInput(dateStr) {
    if (!dateStr || dateStr === '-') return '';
    // YYYY/MM/DD形式をYYYY-MM-DD形式に変換
    return dateStr.replace(/\//g, '-');
}

// モーダルを閉じる
function closeModal() {
    document.getElementById('editModal').style.display = 'none';
}

// フォーム送信処理
async function handleFormSubmit(event) {
    event.preventDefault();
    
    const index = document.getElementById('editIndex').value;
    const zairyuShikaku = document.getElementById('edit在留資格').value;
    const kiManryoDays = document.getElementById('edit既満了日数').value;
    
    // 特定技能1号の場合、既満了日数の入力チェック
    if (zairyuShikaku.includes('特定技能') && (zairyuShikaku.includes('1号') || zairyuShikaku.includes('１号'))) {
        if (kiManryoDays === '' || kiManryoDays === null) {
            alert('特定技能1号の場合、既満了日数の入力は必須です。\n0以上の数値を入力してください。');
            document.getElementById('edit既満了日数').focus();
            return;
        }
        const kiManryoNum = parseInt(kiManryoDays);
        if (isNaN(kiManryoNum) || kiManryoNum < 0) {
            alert('既満了日数は0以上の数値を入力してください。');
            document.getElementById('edit既満了日数').focus();
            return;
        }
    }
    
    const data = {
        '担当者コード': document.getElementById('edit担当者コード').value,
        '氏名１': document.getElementById('edit氏名１').value,
        '氏名２': document.getElementById('edit氏名２').value,
        '在留資格': zairyuShikaku,
        '国籍': document.getElementById('edit国籍').value,
        '在留カード番号': document.getElementById('edit在留カード番号').value,
        '生年月日': document.getElementById('edit生年月日').value,
        '期生': document.getElementById('edit期生').value,
        '許可年月日': document.getElementById('edit許可年月日').value,
        '満了年月日': document.getElementById('edit満了年月日').value,
        '既満了日数': kiManryoDays,
        '設定期限1': document.getElementById('edit設定期限1').value,
        '設定期限2': document.getElementById('edit設定期限2').value,
        '設定期限3': document.getElementById('edit設定期限3').value,
    };
    
    try {
        let response;
        if (index === '') {
            // 新規追加
            updateStatus('データを追加中...');
            response = await fetch('/api/data/add', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(data)
            });
        } else {
            // 更新
            updateStatus('データを更新中...');
            response = await fetch(`/api/data/update/${index}`, {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(data)
            });
        }
        
        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.error || '保存に失敗しました');
        }
        
        const result = await response.json();
        updateStatus(result.message);
        closeModal();

        // データを再読み込み
        await loadData();
        await loadSummary();
        markUnsavedChanges();
    } catch (error) {
        console.error('Error saving data:', error);
        updateStatus('エラー: ' + error.message);
        alert('保存に失敗しました:\n' + error.message);
    }
}

// データを削除
async function deleteRow(index) {
    if (!confirm('このデータを削除しますか？')) {
        return;
    }

    const finalConfirmMessage = '最終確認: この操作は元に戻せません。\n本当に削除しますか？';
    if (!confirm(finalConfirmMessage)) {
        return;
    }
    
    try {
        updateStatus('データを削除中...');
        const response = await fetch(`/api/data/delete/${index}`, {
            method: 'DELETE'
        });
        
        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.error || '削除に失敗しました');
        }
        
        const result = await response.json();
        updateStatus(result.message);
        
        // データを再読み込み
        await loadData();
        await loadSummary();
        markUnsavedChanges();
    } catch (error) {
        console.error('Error deleting data:', error);
        updateStatus('エラー: ' + error.message);
        alert('削除に失敗しました:\n' + error.message);
    }
}

function extractFilenameFromDisposition(disposition) {
    const filenameRegex = /filename\*=UTF-8''([^;]+)|filename="?([^";]+)"?/i;
    const matches = filenameRegex.exec(disposition);
    if (!matches) return null;
    const encoded = matches[1];
    const plain = matches[2];
    if (encoded) {
        try {
            return decodeURIComponent(encoded);
        } catch (e) {
            return encoded;
        }
    }
    return plain;
}

function parseSummaryTimestamp(timestamp) {
    if (!timestamp) return null;
    const normalized = timestamp.replace(' ', 'T');
    const parsed = new Date(normalized);
    return isNaN(parsed.getTime()) ? null : parsed;
}

function updateLastExportInfo(summary) {
    const lastExportElement = document.getElementById('lastExportInfo');
    if (!lastExportElement) return;

    if (summary.last_exported_at) {
        const filenameSuffix = summary.last_export_filename ? ` (${summary.last_export_filename})` : '';
        lastExportElement.textContent = `最終保存: ${summary.last_exported_at}${filenameSuffix}`;
    } else {
        lastExportElement.textContent = '最終保存: 未保存';
    }

    lastExportedAt = parseSummaryTimestamp(summary.last_exported_at);

    if (lastChangeTimestamp && lastExportedAt && lastExportedAt.getTime() >= lastChangeTimestamp) {
        hasUnsavedChanges = false;
        lastChangeTimestamp = null;
    }

    updateUnsavedIndicator();
}

function markUnsavedChanges() {
    hasUnsavedChanges = true;
    lastChangeTimestamp = Date.now();
    updateUnsavedIndicator();
}

function resetUnsavedChanges() {
    hasUnsavedChanges = false;
    lastChangeTimestamp = null;
    updateUnsavedIndicator();
}

function updateUnsavedIndicator() {
    const indicator = document.getElementById('unsavedChangesIndicator');
    const exportButton = document.getElementById('processedExportButton');

    if (!indicator || !exportButton) {
        return;
    }

    if (hasUnsavedChanges) {
        indicator.style.display = 'block';
        exportButton.classList.add('btn-unsaved');
        exportButton.setAttribute('title', '未保存の変更があります。処理済みデータ保存を実行してください。');
    } else {
        indicator.style.display = 'none';
        exportButton.classList.remove('btn-unsaved');
        exportButton.removeAttribute('title');
    }
}

function getDeadlineLevelLimits(row) {
    // レスポンスに期限レベル*_上限が含まれていれば優先
    const level1 = toNumberOrNull(row['期限レベル1_上限']);
    const level2 = toNumberOrNull(row['期限レベル2_上限']);
    const level3 = toNumberOrNull(row['期限レベル3_上限']);

    return { level1, level2, level3 };
}

function toNumberOrNull(value) {
    if (value === undefined || value === null || value === '') {
        return null;
    }
    const num = Number(value);
    return Number.isFinite(num) ? Math.trunc(num) : null;
}
