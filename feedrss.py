import feedparser
import certifi
import ssl
import urllib.request
import pandas as pd
from bs4 import BeautifulSoup

def fetch_news_from_rss(rss_url, asset):
    """Fetch news items from an RSS feed URL and tag with the asset name."""
    try:
        context = ssl.create_default_context(cafile=certifi.where())
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
        req = urllib.request.Request(rss_url, headers=headers)
        response = urllib.request.urlopen(req, context=context)
        feed = feedparser.parse(response)
        if feed.bozo:
            print(f"Failed to parse feed: {rss_url} (Error: {feed.bozo_exception})")
            return []
        news_items = []
        for entry in feed.entries:
            summary = entry.get('summary', 'No summary available')
            # Clean the summary from HTML tags
            clean_summary = BeautifulSoup(summary, "html.parser").get_text()
            news_item = {
                'Title': entry.title,
                'Link': entry.link,
                'Published': entry.get('published', 'No date available'),
                'Summary': clean_summary,
                'Asset': asset  # Tagging with the asset name
            }
            news_items.append(news_item)
        return news_items
    except Exception as e:
        print(f"Error fetching news from {rss_url}: {e}")
        return []

def display_news(news_items):
    """Display news items with their asset tags."""
    if not news_items:
        print("No news items found.")
        return
    for item in news_items:
        print(f"Asset: {item['Asset']}")
        print(f"Title: {item['Title']}")
        print(f"Link: {item['Link']}")
        print(f"Published: {item['Published']}")
        print(f"Summary: {item['Summary']}\n")
    print("-" * 80)

def save_news_to_excel(news_items, filename):
    """Save news items to an Excel file."""
    try:
        df = pd.DataFrame(news_items)
        df.to_excel(filename, index=False, engine='openpyxl')
        print(f"News items successfully saved to {filename}")
    except Exception as e:
        print(f"Error saving news items to {filename}: {e}")

def read_assets_from_file(filename):
    """Read the list of assets from a file."""
    try:
        with open(filename, 'r') as file:
            assets = [line.strip() for line in file if line.strip()]
        return assets
    except Exception as e:
        print(f"Error reading assets from file {filename}: {e}")
        return []

def read_feeds_from_file(filename):
    """Read the list of RSS feed URLs and their corresponding assets from a file."""
    try:
        feeds = []
        with open(filename, 'r') as file:
            for line in file:
                parts = line.strip().split(maxsplit=1)
                if len(parts) == 2:
                    url, asset = parts
                    feeds.append((url, asset))
        return feeds
    except Exception as e:
        print(f"Error reading feeds from file {filename}: {e}")
        return []

if __name__ == "__main__":
    # Read the list of assets from asset.txt
    assets = read_assets_from_file('asset.txt')

    # Read the list of RSS feed URLs and their corresponding assets from feeds.txt
    feeds = read_feeds_from_file('feeds.txt')

    # Filter feeds based on the assets listed in asset.txt
    filtered_feeds = [feed for feed in feeds if feed[1] in assets]

    # Fetch and display news from the filtered feeds
    all_news_items = []
    for rss_url, asset in filtered_feeds:
        print(f"Fetching news from: {rss_url}")
        news_items = fetch_news_from_rss(rss_url, asset)
        display_news(news_items)
        all_news_items.extend(news_items)
        print("=" * 80 + "\n")

    # Save the news items to an Excel file
    save_news_to_excel(all_news_items, 'news_items.xlsx')
