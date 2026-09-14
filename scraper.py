import time
import re
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

def get_driver():
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

def get_all_player_urls(driver):
    print("[+] Step 1: Collecting player profile links...")
    driver.get("https://crickslab.com/match-central/players")
    time.sleep(8)
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(3)
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    links = soup.find_all('a', href=True)
    urls = []
    for link in links:
        href = link['href']
        if '/player-details/' in href:
            full_url = href if href.startswith('http') else "https://crickslab.com" + href
            if full_url not in urls:
                urls.append(full_url)
    print(f"[+] Found {len(urls)} player profiles.")
    return urls

def parse_player_profile(driver, url):
    driver.get(url)
    time.sleep(4)
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    text_content = soup.get_text(separator=' ')
    name_tag = soup.find(['h1', 'h2', 'h3'])
    player_name = name_tag.text.strip() if name_tag else "N/A"
    
    def extract_val(pattern, text, default=0):
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            val = m.group(1).replace(',', '').strip()
            return float(val) if '.' in val else int(val)
        return default

    def extract_str(pattern, text, default="N/A"):
        m = re.search(pattern, text, re.IGNORECASE)
        return m.group(1).strip() if m else default

    role = extract_str(r'Role\s*:\s*([A-Za-z\s]+)', text_content, "All Rounder")
    nationality = extract_str(r'Nationality\s*:\s*([A-Za-z\s]+)', text_content, "N/A")
    city = extract_str(r'City\s*:\s*([A-Za-z\s]+)', text_content, "N/A")
    batting_style = extract_str(r'Batting Style\s*:\s*([A-Za-z\s\-]+)', text_content, "Right Handed")
    bowling_style = extract_str(r'Bowling Style\s*:\s*([A-Za-z\s\-]+)', text_content, "N/A")
    matches = extract_val(r'Matches\s*:?\s*(\d+)', text_content, 0)
    runs = extract_val(r'Runs\s*:?\s*(\d+)', text_content, 0)
    wickets = extract_val(r'Wickets\s*:?\s*(\d+)', text_content, 0)
    bat_m = extract_val(r'Batting Matches\s*:?\s*(\d+)', text_content, matches)
    bat_i = extract_val(r'Innings\s*:?\s*(\d+)', text_content, matches)
    hs = extract_val(r'High Score\s*:?\s*(\d+)', text_content, 0)
    bat_avg = extract_val(r'Batting Avg\s*:?\s*([\d\.]+)', text_content, 0.0)
    bat_sr = extract_val(r'Strike Rate\s*:?\s*([\d\.]+)', text_content, 0.0)
    fifties = extract_val(r'50s\s*:?\s*(\d+)', text_content, 0)
    hundreds = extract_val(r'100s\s*:?\s*(\d+)', text_content, 0)
    fours = extract_val(r'4s\s*:?\s*(\d+)', text_content, 0)
    sixes = extract_val(r'6s\s*:?\s*(\d+)', text_content, 0)
    bowl_m = extract_val(r'Bowling Matches\s*:?\s*(\d+)', text_content, matches)
    bowl_i = extract_val(r'Bowling Innings\s*:?\s*(\d+)', text_content, 0)
    overs = extract_val(r'Overs\s*:?\s*([\d\.]+)', text_content, 0.0)
    maidens = extract_val(r'Maidens\s*:?\s*(\d+)', text_content, 0)
    bowl_runs = extract_val(r'Bowl Runs\s*:?\s*(\d+)', text_content, 0)
    econ = extract_val(r'Economy\s*:?\s*([\d\.]+)', text_content, 0.0)
    bowl_avg = extract_val(r'Bowling Avg\s*:?\s*([\d\.]+)', text_content, 0.0)
    bowl_sr = extract_val(r'Bowling SR\s*:?\s*([\d\.]+)', text_content, 0.0)
    five_w = extract_val(r'5W\s*:?\s*(\d+)', text_content, 0)
    catches = extract_val(r'Catches\s*:?\s*(\d+)', text_content, 0)
    run_outs = extract_val(r'Run Outs\s*:?\s*(\d+)', text_content, 0)

    return {
        'Player Name': player_name,
        'Role (Site)': role,
        'Nationality': nationality,
        'City': city,
        'Batting Style': batting_style,
        'Bowling Style': bowling_style,
        'Career Matches': matches,
        'Career Runs': runs,
        'Career Wickets': wickets,
        'Bat M': bat_m,
        'Bat I': bat_i,
        'Bat Runs': runs,
        'HS': hs,
        'Bat Avg': bat_avg,
        'Bat SR': bat_sr,
        '50s': fifties,
        '100s': hundreds,
        '4s': fours,
        '6s': sixes,
        'Bowl M': bowl_m,
        'Bowl I': bowl_i,
        'Overs': overs,
        'Maidens': maidens,
        'Bowl Runs': bowl_runs,
        'Wkts': wickets,
        'Econ': econ,
        'Bowl Avg': bowl_avg,
        'Bowl SR': bowl_sr,
        '5W': five_w,
        'Catches': catches,
        'Run Outs': run_outs,
        'Profile URL': url
    }

def main():
    driver = get_driver()
    try:
        urls = get_all_player_urls(driver)
        all_players_data = []
        print("[+] Step 2: Scraping detailed career statistics...")
        for idx, url in enumerate(urls, 1):
            print(f"    - Processing ({idx}/{len(urls)}): {url}")
            data = parse_player_profile(driver, url)
            all_players_data.append(data)
            
        df = pd.DataFrame(all_players_data)
        output_filename = "Crickslab_Players_Data.xlsx"
        print("[+] Step 3: Exporting data to formatted multi-sheet Excel file...")
        with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=f'All Players ({len(df)})', index=False)
            batsmen_df = df[df['Role (Site)'].str.contains('Batter|Batting', case=False, na=False)]
            bowlers_df = df[df['Role (Site)'].str.contains('Bowler|Bowling', case=False, na=False)]
            all_rounders_df = df[df['Role (Site)'].str.contains('All Rounder', case=False, na=False)]
            batsmen_df.to_excel(writer, sheet_name='Batsmen', index=False)
            bowlers_df.to_excel(writer, sheet_name='Bowlers', index=False)
            all_rounders_df.to_excel(writer, sheet_name='All Rounders', index=False)
        print(f"\n[SUCCESS] Completed! Output file saved to '{output_filename}'.")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
