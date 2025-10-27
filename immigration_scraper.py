"""
出入国在留管理庁の情報をスクレイピングするモジュール
"""
import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
import os
from pathlib import Path
import time
from typing import List, Dict, Optional
import logging

# ロギング設定
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class ImmigrationScraper:
    """出入国在留管理庁の情報をスクレイピングするクラス"""
    
    BASE_URL = "https://www.moj.go.jp/"
    
    def __init__(self, output_dir: str = "data"):
        """
        初期化
        
        Args:
            output_dir: 出力ディレクトリ
        """
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        })
    
    def fetch_news(self, category: str) -> List[Dict]:
        """
        指定されたカテゴリのニュースを取得
        
        Args:
            category: カテゴリ（'zairyu'（在留関係）など）
            
        Returns:
            List[Dict]: ニュースアイテムのリスト
        """
        url = f"{self.BASE_URL}nyuukokukanri/kouhou/nyuukokukanri{category}.html"
        logger.info(f"ニュースを取得中: {url}")
        
        try:
            response = self.session.get(url, timeout=10)
            response.raise_for_status()
            
            soup = BeautifulSoup(response.content, 'html.parser')
            news_items = []
            
            # ニュースアイテムを抽出（実際のHTML構造に応じて調整が必要）
            for item in soup.select('.newsList li'):
                date_elem = item.select_one('.date')
                title_elem = item.select_one('a')
                
                if date_elem and title_elem:
                    news_items.append({
                        'date': date_elem.text.strip(),
                        'title': title_elem.text.strip(),
                        'url': self.BASE_URL.rstrip('/') + title_elem['href'].lstrip('./')
                    })
            
            return news_items
            
        except Exception as e:
            logger.error(f"ニュースの取得中にエラーが発生しました: {e}")
            return []
    
    def filter_news_by_keywords(self, news_items: List[Dict], keywords: List[str]) -> List[Dict]:
        """
        キーワードでニュースをフィルタリング
        
        Args:
            news_items: ニュースアイテムのリスト
            keywords: 検索キーワードのリスト
            
        Returns:
            List[Dict]: フィルタリングされたニュースアイテム
        """
        if not news_items or not keywords:
            return []
            
        filtered = []
        for item in news_items:
            title = item.get('title', '').lower()
            if any(keyword.lower() in title for keyword in keywords):
                filtered.append(item)
                
        return filtered
    
    def save_to_csv(self, data: List[Dict], filename: str) -> str:
        """
        データをCSVに保存
        
        Args:
            data: 保存するデータ
            filename: ファイル名（拡張子なし）
            
        Returns:
            str: 保存されたファイルのパス
        """
        if not data:
            return ""
            
        # 日付をファイル名に追加
        today = datetime.now().strftime('%Y%m%d')
        output_path = self.output_dir / f"{filename}_{today}.csv"
        
        df = pd.DataFrame(data)
        df.to_csv(output_path, index=False, encoding='utf-8-sig')
        
        logger.info(f"データを保存しました: {output_path}")
        return str(output_path)
    
    def scrape_immigration_updates(self):
        """
        在留資格関連の更新情報をスクレイピング
        """
        logger.info("在留資格関連の更新情報を取得中...")
        
        # 在留関係のニュースを取得
        news_items = self.fetch_news('zairyu')
        
        # 技能実習と特定技能に関連するキーワードでフィルタリング
        keywords = ['技能実習', '特定技能', '在留資格', '更新', '変更']
        filtered_news = self.filter_news_by_keywords(news_items, keywords)
        
        # 結果を保存
        if filtered_news:
            output_file = self.save_to_csv(
                filtered_news, 
                'immigration_status_updates'
            )
            logger.info(f"更新情報を {len(filtered_news)} 件見つけました: {output_file}")
        else:
            logger.info("新しい更新情報は見つかりませんでした。")
        
        return filtered_news


def main():
    """メイン処理"""
    try:
        # スクレイパーの初期化
        scraper = ImmigrationScraper()
        
        # 在留資格関連の更新情報を取得して保存
        updates = scraper.scrape_immigration_updates()
        
        if updates:
            print("\n=== 取得した更新情報 ===")
            for i, item in enumerate(updates, 1):
                print(f"{i}. [{item['date']}] {item['title']}")
                print(f"   URL: {item['url']}\n")
        else:
            print("新しい更新情報は見つかりませんでした。")
            
    except Exception as e:
        logger.error(f"エラーが発生しました: {e}", exc_info=True)


if __name__ == "__main__":
    main()
