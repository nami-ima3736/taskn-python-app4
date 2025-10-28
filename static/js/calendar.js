// 在留資格管理システム - カレンダー JavaScript

// グローバル変数
let calendarData = {};
let currentYear = new Date().getFullYear();
let currentMonth = new Date().getMonth() + 1; // 1-12

// ページ読み込み時の初期化
document.addEventListener('DOMContentLoaded', function() {
    // ファイル選択イベント
    document.getElementById('fileInput').addEventListener('change', handleFileUpload);
    
    // カレンダー初期化
    initializeCalendar();
    loadSummary();
});

// カレンダー初期化
function initializeCalendar() {
    updateMonthYearDisplay();
    loadCalendarData();
}

// 月年表示を更新
function updateMonthYearDisplay() {
    const monthNames = ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'];
    document.getElementById('currentMonthYear').textContent = `${currentYear}年${monthNames[currentMonth - 1]}`;
}

// カレンダーデータを読み込む
async function loadCalendarData() {
    try {
        updateStatus('カレンダーデータを読み込み中...');
        const response = await fetch(`/api/calendar?year=${currentYear}&month=${currentMonth}`);
        
        if (!response.ok) {
            throw new Error('カレンダーデータの読み込みに失敗しました');
        }
        
        const result = await response.json();
        calendarData = result.calendar_data;
        
        generateCalendar();
        updateStatus(`カレンダーを更新しました (${currentYear}年${currentMonth}月)`);
    } catch (error) {
        console.error('Error loading calendar data:', error);
        updateStatus('エラー: ' + error.message);
        generateCalendar(); // データなしでカレンダー生成
    }
}

// カレンダーを生成
function generateCalendar() {
    const tbody = document.getElementById('calendarBody');
    tbody.innerHTML = '';
    
    const monthNames = ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'];
    const daysInMonth = new Date(currentYear, currentMonth, 0).getDate();
    const firstDay = new Date(currentYear, currentMonth - 1, 1).getDay();
    
    let day = 1;
    for (let i = 0; i < 6; i++) { // 最大6週間
        const row = document.createElement('tr');
        
        for (let j = 0; j < 7; j++) { // 7日
            const cell = document.createElement('td');
            cell.className = 'calendar-cell';
            
            if (i === 0 && j < firstDay) {
                // 前の月の日付
                cell.classList.add('other-month');
            } else if (day > daysInMonth) {
                // 次の月の日付
                cell.classList.add('other-month');
                day = 1; // 次の月用
            } else {
                // 現在の月の日付
                const dateStr = `${currentYear}-${String(currentMonth).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
                const today = new Date();
                const todayStr = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
                
                if (dateStr === todayStr) {
                    cell.classList.add('today');
                }
                
                if (calendarData[dateStr]) {
                    cell.classList.add('has-data');
                    // データ表示
                    const data = calendarData[dateStr];
                    let content = `<div class="date-number">${day}</div>`;
                    
                    ['deadline1', 'deadline2', 'deadline3'].forEach(type => {
                        if (data[type].length > 0) {
                            content += `<div class="deadline-${type.slice(-1)}">`;
                            data[type].forEach(item => {
                                content += `<div class="person-item ${type}">`;
                                content += `<span class="code">${item['担当者コード']}</span>`;
                                content += `<span class="name">${item['氏名２']}</span>`;
                                content += `<span class="status">${item['在留資格']}</span>`;
                                content += `</div>`;
                            });
                            content += `</div>`;
                        }
                    });
                    
                    cell.innerHTML = content;
                } else {
                    cell.innerHTML = `<div class="date-number">${day}</div>`;
                }
                
                day++;
            }
            
            row.appendChild(cell);
        }
        
        tbody.appendChild(row);
        
        // すべての日付を処理したら終了
        if (day > daysInMonth) {
            break;
        }
    }
}

// 月ナビゲーション
function navigateMonth(direction) {
    currentMonth += direction;
    if (currentMonth < 1) {
        currentMonth = 12;
        currentYear--;
    } else if (currentMonth > 12) {
        currentMonth = 1;
        currentYear++;
    }
    updateMonthYearDisplay();
    loadCalendarData();
}

// 年ナビゲーション
function navigateYear(direction) {
    currentYear += direction;
    updateMonthYearDisplay();
    loadCalendarData();
}

// 今日に移動
function goToToday() {
    const today = new Date();
    currentYear = today.getFullYear();
    currentMonth = today.getMonth() + 1;
    updateMonthYearDisplay();
    loadCalendarData();
}

// カレンダーを更新
function refreshCalendar() {
    loadCalendarData();
}

// ステータスを更新
function updateStatus(message) {
    document.getElementById('statusBar').textContent = message;
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
    } catch (error) {
        console.error('Error loading summary:', error);
    }
}

// ファイルアップロード処理
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
        
        const result = await response.json();
        
        if (result.success) {
            updateStatus('ファイルのアップロードが完了しました');
            loadSummary();
            loadCalendarData();
        } else {
            updateStatus('エラー: ' + result.error);
        }
    } catch (error) {
        console.error('Error uploading file:', error);
        updateStatus('アップロードエラー: ' + error.message);
    }
}

// アラートリスト出力
function exportAlertList() {
    window.open('/api/export/alert', '_blank');
}

// 処理済みデータ保存
function exportProcessedData() {
    window.open('/api/export/processed', '_blank');
}
