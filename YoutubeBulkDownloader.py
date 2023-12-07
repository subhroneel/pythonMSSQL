from googleapiclient.discovery import build
from tkinter.ttk import *
from tkinter import *
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import subprocess
from bs4 import BeautifulSoup
import os

class YoutubeBukDownloader:
    def __init__(self,root):
        self.root = root
        self.root.title("Social Media Analyser")
        # Initialize Pygame mixer
        self.root.geometry('550x400')
        self.mediaLink_var = StringVar()
        self.search_str = StringVar()
        self.videoLinks: dict = {}
        cmbMediaNameFrame = LabelFrame(self.root, text="Enter Channel or Playlist", font=("Arial", 9, "bold"), bg="#4d65a3",
                              fg="white", bd=1, relief=GROOVE)
        cmbMediaNameFrame.grid(row=0, column=0, padx=0)
        self.mediaLink = Entry(cmbMediaNameFrame, textvariable=self.mediaLink_var, state="normal", width=40)
        self.mediaLink.insert('end', "https://www.youtube.com/@SubhroneelGanguly")
        self.mediaLink.config(state='normal')
        self.mediaLink.grid(row=1, column=1, padx=0)
        self.btnGetStats = Button(cmbMediaNameFrame, text="List All", command=self.get_links)
        self.btnGetStats.grid(row=1, column=2, padx=2)
        cmbSearchNameFrame = LabelFrame(self.root, text="Filter List", font=("Arial", 9, "bold"), bg="#4d65a3",
                              fg="white", bd=1, relief=GROOVE)
        cmbSearchNameFrame.grid(row=1, column=0, padx=0)
        self.searchBox = Entry(cmbSearchNameFrame, textvariable=self.search_str, state="normal", width=40)
        self.searchBox.insert('end', "")
        self.searchBox.config(state='normal')
        self.searchBox.bind('<KeyPress>', self.cb_search)
        self.searchBox.grid(row=2, column=1, padx=0)

        tblListframe = LabelFrame(self.root, text="List of Videos", font=("arial", 9, "bold"), bg="#8F00FF",
                                  fg="white",
                                  bd=1, relief=GROOVE)
        tblListframe.grid(row=2, column=0, padx=0)
        # Create UI elements
        scroll_y = Scrollbar(tblListframe, orient=VERTICAL)
        scroll_x = Scrollbar(tblListframe, orient=HORIZONTAL)
        self.LB = Listbox(tblListframe, yscrollcommand=scroll_y.set, selectbackground="#8d8df6", selectmode=EXTENDED,
                          width=70,
                          font=("arial", 9, "bold"), bg="#c9f56f", fg="navyblue", bd=5, relief=GROOVE)
        scroll_y.config(command=self.LB.yview)
        scroll_x.config(command=self.LB.xview)
        scroll_y.pack(side=RIGHT, fill=Y)
        scroll_x.pack(side=BOTTOM, fill=X)
        self.LB.bind('<<ListboxSelect>>', self.items_selected)
        self.LB.pack(fill=BOTH)
        self.LB['state'] = DISABLED
        downLoadframe = LabelFrame(self.root, text="Select Type", font=("arial", 9, "bold"), bg="#8F00FF",
                                  fg="white",
                                  bd=1, relief=GROOVE)
        downLoadframe.grid(row=3, column=0, padx=0)
        self.cmbType = Combobox(downLoadframe, width=27, textvariable=StringVar(),
                              values=['Video-640x360-360p','Video-1280x720-720p','Audio-mp4a.40.5-ultralow','Audio-opus-ultralow','Audio-mp4a.40.5-low','Audio-mp4a.40.5-low'])
        self.cmbType.set("")
        self.cmbType.bind('<<ComboboxSelected>>', self.oncmbTypeSelect)
        self.cmbType.grid(row=0, column=0, padx=2)

        self.btnDownloadVideos = Button(downLoadframe, text="Download", command=self.downLoadVideo)
        self.btnDownloadVideos.grid(row=0, column=1, padx=2)
        self.btnDownloadVideos['state'] = DISABLED

        pbframe = LabelFrame(self.root, text="", font=("arial", 12, "bold"), bg="#8F00FF", fg="white",
                             bd=5, relief=GROOVE)
        pbframe.grid(row=4, column=0, padx=0)
        self.pb = Progressbar(pbframe, orient='horizontal', mode='determinate', length=480)
        # self.pb.grid(column=0, row=0, columnspan=2, padx=20, pady=40)
        self.pb.pack()

    def oncmbTypeSelect(self, event):

        # self.LB.configure(exportselection=False)
        if self.cmbType.get():
            self.btnDownloadVideos['state'] = NORMAL
        self.LB.focus_set()


    def downLoadVideo(self):
        self.pb['maximum'] = len(self.videoList)
        self.pb['value'] = 0
        for link in self.videoList:
            # print(link)
            os.system('yt-dlp -f ' + self.getDownloadType(self.cmbType.get()) + ' ' + link)
            self.pb['value'] += 1
            self.root.update_idletasks()
    
    def items_selected(self, event):
        # get selected indices
        selected_indices = self.LB.curselection()
        w = event.widget
        # get selected items
        selected_langs = ",".join([self.videoLinks[self.LB.get(i)] for i in selected_indices])
        self.videoList = selected_langs.split(',')
        if len(self.videoList) > 0:
            self.cmbType['state'] = NORMAL
        else:
            self.btnDownloadVideos['state'] = DISABLED
            self.cmbType['state'] = DISABLED

    def setcmbMediaNameText(self, event):
        self.MediaName = self.cmbMediaName.get()

    def cb_search(self, event):
        
        sstr = self.search_str.get()
        self.LB.delete(0, END)
        if sstr == "":
            self.fill_listbox(self.videoLinks)   
            self.root.update_idletasks()
            return
        # If filter removed show all data
    
        filtered_data = list()
        for key in self.videoLinks:
            if key.find(sstr) >= 0:
                filtered_data.append(key)
    
        self.fill_listbox(filtered_data)   
        self.root.update_idletasks()
    
    def fill_listbox(self,ld):
        for item in ld:
            self.LB.insert(END, item)

    def initiate_chrome_options(self):
        # Set up Chrome options for headless mode
        chrome_options = Options()
        chrome_options.add_argument('--headless')  # Run Chrome in headless mode
        chrome_options.add_argument('--disable-gpu')  # Disable GPU acceleration in headless mode

        # Initialize the WebDriver with Chrome options
        self.driver = webdriver.Chrome(options=chrome_options)

    def check_connection(self):
        try:
            # Navigate to the URL
            self.driver.get(self.mediaLink_var.get())

            # Wait for the specific span element with the class to be present in the DOM
            content = EC.presence_of_element_located((By.CSS_SELECTOR, 'body'))
            WebDriverWait(self.driver, 20).until(content)
        except TimeoutException:
            print("Timed out waiting for the target element to be present on the page.")
        finally:
            # driver.quit()
            pass    

    def get_links(self):
        self.initiate_chrome_options()
        self.check_connection()
        content = self.driver.page_source.encode('utf-8').strip()
        beautiful_soup = BeautifulSoup(content, "html.parser")
        videoTitles = beautiful_soup.findAll('a', {'id':'video-title'})
        self.videoLinks = {}
        self.LB['state'] = NORMAL   
        for link in videoTitles:
            self.LB.insert(END, link['title'])
            self.videoLinks.update({link['title']: 'https://www.youtube.com/' + link['href']})
        self.root.update_idletasks()

    def get_all_links(self):
        self.initiate_chrome_options()
        self.check_connection()
        content = self.driver.page_source.encode('utf-8').strip()
        beautiful_soup = BeautifulSoup(content, "html.parser")
        channel_Id = beautiful_soup.find('meta', {'itemprop':'identifier'})        
        # Replace 'YOUR_API_KEY' with the actual API key
        API_KEY = 'AIzaSyDNZXqQV8tgQcHyb4ArJZPzOpfPkEx01Y0'
        youtube = build('youtube', 'v3', developerKey=API_KEY)

        # Replace 'CHANNEL_ID' with the ID of the YouTube channel
        CHANNEL_ID = channel_Id['content']

        # Get the list of videos from the channel
        next_page_token = None
        self.videoLinks = {}
        self.LB['state'] = NORMAL   
        while True:
            request = youtube.search().list(
                part='snippet',
                channelId=CHANNEL_ID,
                order='date',
                type='video',
                maxResults=50,  # Adjust as needed
                pageToken=next_page_token
            )
            response = request.execute()

            # Print video titles
            for item in response['items']:
                video_title = item['snippet']['title']
                video_id = item['id']['videoId']
                video_url = f'https://www.youtube.com/watch?v={video_id}'
                self.LB.insert(END, video_title)
                self.root.update_idletasks()
                self.videoLinks.update({video_title: video_url})
                # print(f'{video_title}: {video_url}')
            # Check if there are more pages
            next_page_token = response.get('nextPageToken')
            if not next_page_token:
                break



# Now, videos_info array contains information for each available format


    def getDownloadType(self, event) -> str:
        if self.cmbType.get() == 'Audio-mp4a.40.5-ultralow':
            return '599'
        elif self.cmbType.get() == 'Audio-opus-ultralow':
            return '600'
        elif self.cmbType.get() == 'Audio-mp4a.40.5-low':
            return '139'
        elif self.cmbType.get() == 'Audio-opus-low':
            return '250'
        elif self.cmbType.get() == 'Video-640x360-360p':
            return '18'
        elif self.cmbType.get() == 'Video-1280x720-720p':
            return '22'



if __name__ == "__main__":
    root = Tk()
    YoutubeBukDownloader(root)
    root.mainloop()


    # get_twitter_stats()

    # twitter css-1qaijid r-bcqeeo r-qvutc0 r-poiln3 r-1b43r93 r-1cwl3u0 r-b88u0q
    #instagram html-span xdj266r x11i5rnm xat24cr x1mh8g0r xexx8yu x4uap5 x18d9i69 xkhd6sd x1hl2dhg x16tdsg8 x1vvkbs