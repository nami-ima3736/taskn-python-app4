"""
スクレイピングを定期的に実行するスケジューラ
月2回（1日と15日）に自動実行
"""
import schedule
import time
import logging
from datetime import datetime
from pathlib import Path
from immigration_scraper import ImmigrationScraper

# ロギング設定
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("scraper_scheduler.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

def run_scraper():
    """スクレイパーを実行する関数"""
    try:
        logger.info("スクレイピングを開始します...")
        
        # スクレイパーの初期化と実行
        scraper = ImmigrationScraper()
        updates = scraper.scrape_immigration_updates()
        
        if updates:
            logger.info(f"スクレイピングが完了しました。{len(updates)}件の更新情報を取得しました。")
        else:
            logger.info("スクレイピングは完了しましたが、新しい更新情報は見つかりませんでした。")
            
    except Exception as e:
        logger.error(f"スクレイピング中にエラーが発生しました: {e}", exc_info=True)

def schedule_jobs():
    """スケジュールを設定"""
    # 毎月1日と15日の午前9時に実行
    schedule.every().day.at("09:00").do(
        lambda: run_scraper() if datetime.now().day in [1, 15] else None
    )
    
    logger.info("スケジューラを開始しました。月1日と15日の午前9時に自動実行されます。")
    logger.info("Ctrl+C で終了します。")
    
    # 初回実行（テスト用）
    run_scraper()
    
    # メインループ
    while True:
        schedule.run_pending()
        time.sleep(60)  # 1分ごとにチェック

def main():
    try:
        schedule_jobs()
    except KeyboardInterrupt:
        logger.info("スケジューラを終了します。")
    except Exception as e:
        logger.error(f"予期しないエラーが発生しました: {e}", exc_info=True)

if __name__ == "__main__":
    main()
