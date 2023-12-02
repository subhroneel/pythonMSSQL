import os
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from tkinter.constants import VERTICAL, GROOVE, RIGHT, BOTH, DISABLED
from tkinter.ttk import Scrollbar

import pygame


class mediaPlayer:
    def __init__(self, root):
        self.slider_set_flag = None
        self.root = root
        self.root.title("Media Player")
        # Initialize Pygame mixer
        self.root.geometry('650x300')
        self.dir_path_var = StringVar()
        self.songlist: dict = {}
        self.song_length = 0

        pygame.init()
        pygame.mixer.init(channels=2)
        self.root.columnconfigure(0, weight=1)
        tblListframe = LabelFrame(self.root, text="List of Tables", font=("arial", 9, "bold"), bg="#8F00FF",
                                  fg="white",
                                  bd=1, relief=GROOVE)
        tblListframe.grid(row=0, column=0, padx=0)
        # Create UI elements
        scroll_y = Scrollbar(tblListframe, orient=VERTICAL)
        self.LB = Listbox(tblListframe, yscrollcommand=scroll_y.set, selectbackground="#8d8df6", selectmode=SINGLE,
                          width=50,
                          font=("arial", 9, "bold"), bg="#c9f56f", fg="navyblue", bd=5, relief=GROOVE)
        scroll_y.config(command=self.LB.yview)
        scroll_y.pack(side=RIGHT, fill=Y)
        self.LB.bind('<<ListboxSelect>>', self.items_selected)
        self.LB.pack(fill=BOTH)
        self.LB['state'] = DISABLED

        posframe = LabelFrame(self.root, text="", font=("arial", 9, "bold"), bg="#8F00FF",
                              fg="white",
                              bd=1, relief=GROOVE)
        posframe.grid(row=1, column=0, padx=5, sticky="nsew")

        self.position = Scale(posframe, from_=0, to=1000, orient='horizontal', showvalue=False)
        self.position.bind("<B1-Motion>", self.change_position)
        self.position.pack(fill="both", expand=True)
        tblButtonframe = LabelFrame(self.root, text="List of Tables", font=("arial", 9, "bold"), bg="#8F00FF",
                                    fg="white",
                                    bd=1, relief=GROOVE)
        tblButtonframe.grid(row=2, column=0, padx=5)

        self.scale = Scale(tblButtonframe, from_=0, to=200, orient='horizontal')
        self.scale.grid(row=1, column=0, padx=0)
        self.scale.bind("<B1-Motion>", self.change_volume)

        self.open_button = tk.Button(tblButtonframe, text="Open Folder", command=self.open_directory)
        self.open_button.grid(row=1, column=1, padx=2)
        # self.play_button.pack(side="left", padx=5)

        self.play_button = tk.Button(tblButtonframe, text="Play", command=self.play_music)
        self.play_button.grid(row=1, column=2, padx=5)
        # self.play_button.pack(side="left", padx=5)

        self.pause_button = tk.Button(tblButtonframe, text="Pause", command=self.pause_music)
        self.pause_button.grid(row=1, column=3, padx=5)
        # self.pause_button.pack(side="left", padx=5)

        self.prev_button = tk.Button(tblButtonframe, text="Previous", command=self.prev_music)
        self.prev_button.grid(row=1, column=4, padx=5)

        self.next_button = tk.Button(tblButtonframe, text="Next", command=self.next_music)
        self.next_button.grid(row=1, column=5, padx=5)

        self.stop_button = tk.Button(tblButtonframe, text="Stop", command=self.stop_music)
        self.stop_button.grid(row=1, column=6, padx=5)
        # self.stop_button.pack(side="left", padx=5)

        self.choose_file_button = tk.Button(tblButtonframe, text="Choose File", command=self.choose_file)
        self.choose_file_button.grid(row=1, column=7, padx=5)
        self._job = self.root.after(1000, self.update_slider_position)
        # self.choose_file_button.pack(side="left", padx=5)

    def update_slider_position(self):
        # if not self.slider_set_flag:
        # if self._job:
        #     self.root.after_cancel(self._job)
        # print(pygame.mixer.music.get_pos())
        if pygame.mixer.music.get_pos()!= int(pygame.mixer.music.get_pos() / 1000):
            self.position.set(pygame.mixer.music.get_pos() / 1000)
            val = self.getTime(int(self.position.get()))
            self.position.config(label=val)
            self._job = self.root.after(1000, self.update_slider_position)
        # else:
        # self._job = self.root.after(100, self.update_slider_position)

    def getTime(self, val: int) -> str:
        hour = str(int(val / 3600))
        min = str(int((val % 3600) / 60))
        sec = str((val % 3600) % 60)
        return hour.format('00') + ':' + min + ':' + sec

    def change_position(self, event):
        # if self._job:
        self.root.after_cancel(self._job)
        pygame.mixer.music.set_pos(self.position.get())
        val = self.getTime(int(self.position.get()))
        self.position.config(label=val)
        self.position.update()
        self._job = self.root.after(1000, self.update_slider_position)
        # self.slider_set_flag = False

    def change_volume(self, event):
        print('%d',int(self.scale.get()))
        pygame.mixer.music.set_volume(int(self.scale.get()))
        # time.sleep(0.1)
        # self.soundSource.set_volume(self.scale.get())

    def items_selected(self, event):
        # get selected indices
        selected_indices = self.LB.curselection()
        # w = event.widget
        # # get selected items
        selected_langs = ",".join([self.LB.get(i) for i in selected_indices])
        # self.musicfile = selected_langs
        self.musicfile = self.LB.get(selected_indices)
        # if len(self.musicfile) > 0:
        #     self.exportExcelBtn['state'] = NORMAL
        #     self.exportCSVBtn['state'] = NORMAL
        # else:
        #     self.exportExcelBtn['state'] = DISABLED
        #     self.exportCSVBtn['state'] = DISABLED

    def open_directory(self):
        self.LB['state'] = NORMAL
        directory_path = filedialog.askdirectory()
        if directory_path:
            self.dir_path_var.set(directory_path)
            # self.text_var.set(directory_path)
            fileList = os.listdir(self.dir_path_var.get())
            for x in fileList:
                media_file_path = self.dir_path_var.get() + '/' + x
                if self.check_if_media_file(x[-3:].lower()):
                    filename = media_file_path[media_file_path.rfind("/") - len(media_file_path) + 1:]
                    self.LB.insert(END, filename)
                    self.songlist.update({filename: media_file_path})
            self.LB.pack(expand=True, fill=BOTH, side=LEFT)

    def check_if_media_file(self, filex: str) -> bool:
        return filex == 'mp3' or filex == 'wav' or filex == 'flac' or filex == 'aac' or filex == 'm4a' or filex == 'wma' or filex == 'ogg'

    def play_music(self):
        self.musicfile = self.songlist.get(self.LB.get(self.LB.curselection()[0]))
        pygame.mixer.music.load(self.musicfile)
        pygame.mixer.music.play()
        self.soundSource = pygame.mixer.Sound(self.musicfile)
        self.song_length = self.soundSource.get_length()
        self.position.config(to=self.song_length)
        self.position.set(0)
        self.root.update_idletasks()

    def prev_music(self):
        current_index = self.LB.curselection()
        if current_index:
            next_index = (current_index[0] - 1) % len(self.songlist)
            self.LB.selection_clear(0, tk.END)
            self.LB.selection_set(next_index)
            self.LB.activate(next_index)
            self.play_music()

    def next_music(self):
        current_index = self.LB.curselection()
        if current_index:
            next_index = (current_index[0] + 1) % len(self.songlist)
            self.LB.selection_clear(0, tk.END)
            self.LB.selection_set(next_index)
            self.LB.activate(next_index)
            self.play_music()

    def pause_music(self):
        pygame.mixer.music.pause()

    def stop_music(self):
        pygame.mixer.music.stop()

    def choose_file(self):
        self.musicfile = filedialog.askopenfilename(defaultextension=".mp3",
                                                    filetypes=[("MP3 files", "*.mp3"),
                                                               ("WAV files", "*.wav"),
                                                               ("All files", "*.*")])
        self.LB['state'] = NORMAL
        self.LB.insert(END, self.musicfile)


if __name__ == "__main__":
    root = tk.Tk()
    mediaPlayer(root)
    root.mainloop()
