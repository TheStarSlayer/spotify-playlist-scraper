from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

from bs4 import BeautifulSoup, SoupStrainer

import yt_dlp

import time
import pandas as pd
from urllib.parse import quote_plus

import os
from sanitize_filename import sanitize
from concurrent.futures import ThreadPoolExecutor

# Parser for bs4
parser = 'lxml'

options = Options()
options.add_argument('--headless=new')
options.add_argument("--disable-notifications")
options.add_argument("window-size=1331,670")
options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36")

# Setting up Selenium
try:
    driver = webdriver.Chrome(service = ChromeService(ChromeDriverManager().install()), options=options)
except FileNotFoundError as fnfe:
    print('chromedriver.exe not found. Re-download from repo and place it in the same location as script...')
    exit(-1)

wait = WebDriverWait(driver, 20)
action = ActionChains(driver)

def init_playlist():
    playlist_url = str(input("Paste your playlist link: "))
    playlist_url = playlist_url.strip()

    driver.get(playlist_url)

    # # Dismiss alert
    # try:
    #     WebDriverWait(driver, 1).until(
    #         EC.presence_of_element_located((By.XPATH, '/html/html')) # Deliberate mistake to prevent open app alert
    #     )
    # except Exception as e:
    #     driver.execute_script('window.open();')
    #     driver.switch_to.window(driver.window_handles[1])
    #     driver.get(playlist_url)

    # Retrieve playlist name
    playlist_name = ''
    try:
        playlist_name = str(wait.until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[4]/div/div[2]/div[6]/div/div[2]/div[1]/div/main/section/div[1]/div[2]/div[3]/span[2]/span/h1'))
        ).text)
        print(f"Playlist Name: {playlist_name}")
    except Exception as e:
        print(e)

    # Retrieve total number of songs
    total_songs = 0
    try:
        total_songs = int((wait.until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[4]/div/div[2]/div[6]/div/div[2]/div[1]/div/main/section/div[1]/div[2]/div[3]/div/div[2]/span[1]'))
        ).text).split(' ')[0])
        print(f"Total songs: {total_songs}")
    except Exception as e:
        print(e)

    # # Remove header as it is harder to control scrollbar with it
    # driver.execute_script('''
    #     let header = document.querySelector(".facDIsOQo9q7kiWc4jSg");
    #     header.parentNode.removeChild(header);
    #                     ''')
    # time.sleep(1)

    # Retrieve scroll bar element
    scroll_bar = driver.find_element(By.XPATH, '/html/body/div[4]/div/div[2]/div[6]/div/div[2]/div[3]/div/div')

    # Remove footer popup
    try:
        footer_popup = wait.until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="onetrust-close-btn-container"]/button'))
        )
        footer_popup.click()
        print("Clicked popup away!")
    except:
        pass

    #Remove 

    # Remove misleading scrollbar if exists
    driver.execute_script('''
        let rightCard = document.querySelector(".LayoutResizer__resize-bar.LayoutResizer__inline-start");
        if (rightCard)
            rightCard.parentNode.removeChild(rightCard);
                        ''')

    return playlist_name, total_songs, scroll_bar

def retrieve_songs_as_soup(scroll_bar, total_songs):
    current_row = 0
    covered_rows = 0
    complete_read = False

    soup_list = []

    parse_only = SoupStrainer('div', attrs={"class": 'fNzI6FqqYnAGcGmJjQRG'})
    soup = BeautifulSoup(driver.page_source, parser, parse_only=parse_only).contents[0].contents[0].contents[1].contents[1]

    current_row = int(soup.contents[0]['aria-rowindex']) # First song of soup
    covered_rows = int(soup.contents[-1]['aria-rowindex']) # Last song of soup
    
    soup_list.append(soup)
    print(f'Captured Soup: {current_row}, {covered_rows}')

    # Robust location of scroll bar
    scroll_size = scroll_bar.size
    scroll_location = scroll_bar.location['y'] + scroll_size['height']
    print(f'Captured Location: {scroll_location}')

    # Activate scrollbar -- simulate mouse movement
    try:
        active_area = wait.until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[4]/div/div[2]/div[6]/div/div[2]/div[3]/div'))
        )
    except:
        pass
        
    action.move_to_element(active_area).perform()
    for i in range(0, 3):
        action.move_by_offset(0, 20).perform()  # Move by a very small offset
        action.move_by_offset(0, -20).perform()  # Move it back

    scroll_size = scroll_bar.size
    action.move_to_element(scroll_bar).perform()
    action.move_by_offset(0, scroll_size['height']).perform()
    action.click().perform()

    # Logic to capture soup
    while not complete_read:
        while current_row not in range(covered_rows - 10, covered_rows + 2):
            # Scroll and compare current soup's top song to previous soup's last song
            scroll(scroll_bar=scroll_bar, backwards=False)
            try:
                soup = BeautifulSoup(driver.page_source, parser, parse_only=parse_only).contents[0].contents[0].contents[1].contents[1]
                current_row = int(soup.contents[0]['aria-rowindex'])
                temp_covered_rows = int(soup.contents[-1]['aria-rowindex'])
            except KeyError:
                time.sleep(2)
                soup = BeautifulSoup(driver.page_source, parser, parse_only=parse_only).contents[0].contents[0].contents[1].contents[1]
                current_row = int(soup.contents[0]['aria-rowindex'])
                temp_covered_rows = int(soup.contents[-1]['aria-rowindex'])
            
            # Move back if current row is greater than covered rows
            while current_row > covered_rows + 1:
                scroll(scroll_bar=scroll_bar, backwards=True)
                soup = BeautifulSoup(driver.page_source, parser, parse_only=parse_only).contents[0].contents[0].contents[1].contents[1]
                current_row = int(soup.contents[0]['aria-rowindex'])

            scroll_location = scroll_bar.location['y'] + scroll_size['height']

            if current_row in range(covered_rows - 10, covered_rows + 2) or temp_covered_rows >= total_songs + 1 or scroll_location >= 500:
                
                soup_list.append(soup)
                covered_rows = int(soup.contents[-1]['aria-rowindex'])
                print(f'Captured Soup: {current_row}, {covered_rows}')
                print(f'Captured Location: {scroll_location}')
                break

        if covered_rows >= total_songs + 1:
            complete_read = True

    return soup_list

# Dynamic scrolling
def scroll(scroll_bar, backwards):
    scroll_size = scroll_bar.size
    # Move to bit lower than middle of the scrollbar
    action.move_to_element(scroll_bar).perform()
    action.move_by_offset(0, int(int(scroll_size['height'])*0.25)).perform()

    if backwards == False:
        print("Forwards!")
        action.click_and_hold().move_by_offset(0, 10).perform()
    else:
        print("Backwards!")
        action.click_and_hold().move_by_offset(0, -2).perform()

    action.release().perform()
    time.sleep(1)

    scroll_location = scroll_bar.location['y'] + scroll_size['height']
    print(f'Scroll bar location: {scroll_location}')

def soup_to_list(soup_list):
    aria_index = 2
    playlist_detes = []

    print("Retrieving playlist details from soup.....", end=' ')

    # Smoothen the collection of soups by disregarding already covered songs and add them to list
    for soup in soup_list:
        for i in range(0, len(soup.contents)):
            if int(soup.contents[i]['aria-rowindex']) < aria_index:
                continue
            song_detes = {}
            song_detes['Song Name'] = soup.contents[i].contents[0].contents[1].contents[1].contents[0].contents[0].text

            song_detes['Singer'] = soup.contents[i].contents[0].contents[1].contents[1].contents[1].contents[0].text

            if song_detes['Singer'] == 'E':
                song_detes['Singer'] = soup.contents[i].contents[0].contents[1].contents[1].contents[2].contents[0].text

            song_detes['Singer'].strip()

            song_detes['Album'] = soup.contents[i].contents[0].contents[2].contents[0].contents[0].text

            song_detes['Duration'] = soup.contents[i].contents[0].contents[-1].contents[1].text

            aria_index += 1
            playlist_detes.append(song_detes)

    print("Done")
    return playlist_detes

def retrieve_youtube_links(playlist_detes, playlist_name):

    root_url = 'https://music.youtube.com/'

    # Retrieve only anchor values
    parse_only = SoupStrainer('a', attrs={'class': 'yt-simple-endpoint style-scope yt-formatted-string', 'spellcheck': 'false'})

    for song in playlist_detes:
        song['In YT music'] = 'N/A'
        # Skip if song does not exist
        if song['Song Name'] == '':
            print('Skipped')
            continue

        searchable_song = song['Song Name'] + ' - ' + song['Singer']
        search_url = root_url + 'search?q=' + str(quote_plus(searchable_song))
        driver.get(search_url)
        
        # Wait till page is loaded
        try:
            wait.until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="contents"]/ytmusic-card-shelf-renderer/div[2]/div[2]/div[1]/div/div[2]/div[1]/yt-formatted-string'))
            )
        except:
            print('Page took too long to load')
            continue

        soup = BeautifulSoup(driver.page_source, parser, parse_only=parse_only)

        # Search atmost 12 links -- (1 for top result, 3 for songs only, 3 for videos only, remaining as unsee-able circumstances)
        for link in soup.find_all('a', limit=12):
            min_song_name = song['Song Name'].strip().lower().replace(' ', '')
            link_name = link.text.strip().lower().replace(' ', '')

            if min_song_name in link_name and 'watch' in str(link['href']):
                song['In YT music'] = root_url + str(link['href'])
                print(f"{song['Song Name']} : {song['In YT music']}")
                break
            else:
                if '(' in min_song_name or '-' in min_song_name:
                    if '(' in min_song_name:
                        min_song_name = min_song_name.split('(')[0].strip().lower()
            
                    if '-' in min_song_name:
                        min_song_name = min_song_name.split('-')[0].strip().lower()

                    if min_song_name in link_name and 'watch' in str(link['href']):
                        song['In YT music'] = root_url + str(link['href'])
                        print(f"{song['Song Name']} : {song['In YT music']}")
                        break

    driver.quit()
    df = pd.DataFrame.from_dict(playlist_detes)
    df.to_excel('{pn}.xlsx'.format(pn=playlist_name), index=False)
    return df

def download_from_youtube(url):
    
    path = os.getenv('LOCALAPPDATA') + 'Microsoft\\WinGet\\Packages\\Gyan.FFmpeg.Essentials_Microsoft.Winget.Source_8wekyb3d8bbwe\\ffmpeg-8.0.1-essentials_build\\bin\\ffmpeg.exe'

    yt_opts = {
        'format': 'bestaudio/best',
        'outtmpl': '%(title)s.%(ext)s',
        'ffmpeg-location': path,
        'postprocessors': [{
            'key': 'FFmpegExtractAudio',
            'preferredcodec': 'mp3',
            'preferredquality': '192',
        }],
        'noplaylist': True,
        'verbose': 0, # No output (except for errors)
    }

    with yt_dlp.YoutubeDL(yt_opts) as ydl:
        ydl.download(url)

# Get user input for simultaneous downloads and list of URLs
def download_songs(playlist_detes, playlist_name, simul_downloads = 6):
    
    directory_name = sanitize(playlist_name)

    if not os.path.exists(directory_name):
        os.mkdir(directory_name)
        print(f"Created {directory_name}")

    os.chdir(directory_name)
    print(f"Current path: {os.getcwd()}")
    
    all_urls = playlist_detes['In YT music'].tolist()
    urls = [url for url in all_urls if url != 'N/A']

    with ThreadPoolExecutor(max_workers=simul_downloads) as executor:
        executor.map(download_from_youtube, urls)

    print("Songs downloaded successfully!")

# Executor Function
def retrieve_spotify_playlist():
    start_retrieve = time.time()

    playlist_name, total_songs, scroll_bar = init_playlist()
    download = str(input("Do you want to download all the songs? (Y/n): ")).lower()
    soup_list = retrieve_songs_as_soup(scroll_bar=scroll_bar, total_songs=total_songs)
    playlist_detes = soup_to_list(soup_list=soup_list)

    playlist_detes_df = retrieve_youtube_links(playlist_detes=playlist_detes, playlist_name=playlist_name)

    end_retrieve = time.time()

    print(f"Time taken to retrieve YT music links: {int(end_retrieve - start_retrieve)}s")

    if download == 'y':
        start_download = time.time()
        download_songs(playlist_detes_df, playlist_name)
        end_download = time.time()
        print(f"Time taken to download songs: {int(end_download - start_download)}s")
    elif download == 'n':
        print('Saving data and exiting..')
    else:
        print('Invalid response chosen for download. Saving data and exiting..')

if __name__ == '__main__':
    retrieve_spotify_playlist()