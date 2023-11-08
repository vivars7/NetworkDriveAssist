import ctypes
import os
import string
import sys
import tkinter as tk
from tkinter import filedialog, ttk
import win32wnet
import webbrowser

# 권한 상승을 위한 함수
def run_as_admin():
    if ctypes.windll.shell32.IsUserAnAdmin() == 0:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)

# 웹 링크를 여는 함수
def open_web_link(url):
    webbrowser.open(url, new=2)  # 새 창/탭에서 URL 열기

# 사용 가능한 드라이브 문자를 반환하는 함수
def get_available_drive_letters():
    used_drive_letters = set()
    for drive in string.ascii_uppercase:
        if os.path.exists(drive + ':'):
            used_drive_letters.add(drive)
    available_drive_letters = [letter for letter in string.ascii_uppercase if letter not in used_drive_letters]
    return available_drive_letters

# 네트워크 공유폴더 주소의 유효성을 검증하는 함수
def validate_share_path(share_path):
    return os.path.exists(share_path)

# 드라이브 문자 새로고침 기능
def refresh_drive_letters():
    drive_letters = get_available_drive_letters()
    drive_letter_combobox['values'] = drive_letters
    if drive_letters:
        drive_letter_combobox.set(drive_letters[0])
    else:
        drive_letter_combobox.set('')

# 네트워크 드라이브로 연결하는 함수
def connect_network_drive():
    # 연결 시도 전 상태 메시지 초기화
    status_label.config(text="", foreground="black")  # 기본 텍스트 색상으로 초기화

    share_path = share_path_entry.get()
    share_path = share_path.replace('/', '\\')  # 슬래시를 백슬래시로 변환

    drive_letter = drive_letter_var.get() + ':'
    
    if not validate_share_path(share_path):
        status_label.config(text="유효하지 않은 주소입니다.", foreground="red")
        return

    nr = win32wnet.NETRESOURCE()
    nr.dwType = 1
    nr.lpLocalName = drive_letter
    nr.lpRemoteName = share_path

    try:
        # 이미 드라이브에 연결되어 있다면 연결 해제
        if os.path.exists(drive_letter):
            win32wnet.WNetCancelConnection2(drive_letter, 0, True)
        
        # 네트워크 드라이브로 연결
        win32wnet.WNetAddConnection2(nr)
        status_label.config(text=f"연결됨: {share_path} as {drive_letter}", foreground="green")

        # 연결된 드라이브를 탐색기로 열기
        os.startfile(drive_letter)
    except win32wnet.error as e:
        status_label.config(text=f"에러: {e}", foreground="red")

# UI 생성
root = tk.Tk()
root.title("Network Drive Assist")

# 프레임 생성
frame = ttk.Frame(root, padding=10)
frame.grid(column=0, row=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# 서버 상 경로 입력
share_path_label = ttk.Label(frame, text="네트워크 경로: ")
share_path_label.grid(column=0, row=0, sticky=tk.W)

share_path_entry = ttk.Entry(frame)
share_path_entry.grid(column=1, row=0, sticky=(tk.W, tk.E))

browse_button = ttk.Button(frame, text="찾아보기", command=lambda: [share_path_entry.delete(0, tk.END), share_path_entry.insert(0, filedialog.askdirectory())])
browse_button.grid(column=2, row=0, sticky=tk.W)

# 할당할 드라이브 문자 선택
drive_letter_label = ttk.Label(frame, text="드라이브 문자: ")
drive_letter_label.grid(column=0, row=1, sticky=tk.W)

drive_letters = get_available_drive_letters()
drive_letter_var = tk.StringVar(value=drive_letters[0] if drive_letters else '')

drive_letter_combobox = ttk.Combobox(frame, textvariable=drive_letter_var, values=drive_letters)
drive_letter_combobox.grid(column=1, row=1, sticky=(tk.W, tk.E))

# 새로고침 버튼 추가
refresh_button = ttk.Button(frame, text="새로고침", command=refresh_drive_letters)
refresh_button.grid(column=2, row=1, sticky=tk.W)

# 연결 버튼
connect_button = ttk.Button(frame, text="연결", command=connect_network_drive)
connect_button.grid(column=1, row=2, sticky=(tk.W, tk.E))

# 상태바 프레임 생성
statusbar_frame = ttk.Frame(root, padding=10)
statusbar_frame.grid(column=0, row=4, sticky=(tk.W, tk.E), columnspan=3)

# 카피라이트 라벨
status_label = ttk.Label(statusbar_frame, text="")
status_label.pack(side=tk.LEFT)

# URL 라벨 (클릭 가능하게)
url = "https://github.com/vivars7/NetworkDriveAssist"
url_label = ttk.Label(statusbar_frame, text="Powerded by irae", cursor="hand2")
url_label.pack(side=tk.RIGHT)
url_label.bind("<Button-1>", lambda e: open_web_link(url))

# UI 시작
root.mainloop()