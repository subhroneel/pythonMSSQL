from tkinter import filedialog, GROOVE, ttk
from tkinter.ttk import *
from tkinter import *
from pyftpdlib.authorizers import DummyAuthorizer
from pyftpdlib.handlers import FTPHandler
from pyftpdlib.servers import ThreadedFTPServer
import threading
import getpass


class pyFtpServer:

    def __init__(self, root):
        self.ftp_server = None
        self.server_thread = None
        self.root = root
        self.root.title('FTP Server')
        self.root.geometry('300x200')
        self.passwd_var = StringVar()
        # self.conn: Connection = None
        # self.tableList: list

        srcStartServerFrame = LabelFrame(self.root, text="Start Server", font=("Arial", 9, "bold"), bg="#4d65a3",
                                         fg="white", bd=0, relief=GROOVE)
        srcStartServerFrame.place(x=0, y=0, width=230, height=80)
        srcPasswdTxtFrame = LabelFrame(srcStartServerFrame, text="Enter Password", font=("Arial", 9, "bold"),
                                       bg="#4d65a3",
                                       fg="white", bd=1, relief=GROOVE)
        srcPasswdTxtFrame.place(x=0, y=0, width=150, height=50)

        # Create a Textbox to display the selected directory path
        passwdTxt = Entry(srcPasswdTxtFrame, textvariable=self.passwd_var,show="*", state="normal", width=40)
        passwdTxt.pack(pady=5)

        ftpConnectBtnFrame = LabelFrame(srcStartServerFrame, text="", font=("Arial", 9, "bold"), bg="#4d65a3",
                                        fg="white", bd=1, relief=GROOVE)
        ftpConnectBtnFrame.place(x=151, y=8, width=80, height=40)
        self.ftpConnectBtn = Button(ftpConnectBtnFrame, text="Open",
                                    command=lambda: self.configure_and_start_ftp_server(),
                                    width=5, height=0, font=("arial", 9, "bold"), fg="#9ed6e6", bg="#4c4f55")
        self.ftpConnectBtn.grid(row=0, column=0, padx=0, pady=1)
        self.ftpConnectBtn.pack()

        srcStopServerFrame = LabelFrame(self.root, text="Stop Server", font=("Arial", 9, "bold"), bg="#4d65a3",
                                        fg="white", bd=0, relief=GROOVE)
        srcStopServerFrame.place(x=0, y=68, width=80, height=50)

        self.ftpDisConnectBtn = Button(srcStopServerFrame, text="Close",
                                       command=lambda: self.stopServer(),
                                       width=5, height=0, font=("arial", 9, "bold"), fg="#9ed6e6", bg="#4c4f55")
        self.ftpDisConnectBtn.grid(row=0, column=0, padx=0, pady=5)
        self.ftpDisConnectBtn.pack()
        srcStopServerFrame = LabelFrame(self.root, text="Server Status", font=("Arial", 9, "bold"), bg="#4d65a3",
                                        fg="white", bd=0, relief=GROOVE)
        srcStopServerFrame.place(x=0, y=120, width=200, height=50)
        self.status_label = Label(srcStopServerFrame, text="Server Status: Stopped", font=("Arial", 9, "bold"), fg="red")
        self.status_label.pack(pady=5)

    def configure_and_start_ftp_server(self):
        try:
            self.ftpConnectBtn.config(state=DISABLED)
            self.ftpDisConnectBtn.config(state=NORMAL)

            # Start the FTP server in a separate thread
            self.server_thread = threading.Thread(target=self.run_server)
            self.server_thread.start()

        except Exception as e:
            self.status_label.config(text=f"Server Status: Error - {str(e)}", fg="red")
    def run_server(self):
        try:
            self.status_label.config(text=f"Server Status: Started")
            authorizer = DummyAuthorizer()

            # Add a user with a specific home directory

            authorizer.add_user(getpass.getuser(), self.passwd_var.get(),
                                self.getPath().format(getpass.getuser()), perm="elradfmw")

            # Add firewall rules (modify as needed)
            authorizer.add_anonymous("/tmp", perm="elradfmw")

            handler = FTPHandler
            handler.authorizer = authorizer

            # Bind to both IPv4 and IPv6 addresses (0.0.0.0 and :: mean all available interfaces)
            self.ftp_server = ThreadedFTPServer(("0.0.0.0", 2221), handler)
            self.ftp_server.serve_forever()
        except Exception as e:
            self.status_label.config(text=f"Server Status: Error - {str(e)}", fg="red")
    def stopServer(self):
        # Stop the FTP server
        self.status_label.config(text=f"Server Status: Stopped")
        self.ftp_server.close_all()

        # self.ftp_server.stop()
        # self.root.quit()

        # Wait for the server thread to finish
        self.server_thread.join()

        self.ftpConnectBtn.config(state=NORMAL)
        self.ftpDisConnectBtn.config(state=DISABLED)

    def getPath(self):
        if getpass.getuser() == 'root':
            return "/{}/"
        else:
            return "/home/{}/"


if __name__ == "__main__":
    root = Tk()
    pyFtpServer(root)
    root.mainloop()
