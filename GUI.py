import tkinter as tk
from tkinter import ttk, Label, Frame
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo

from CRNStatusChecker import checkCRNStatus

# Create Main Window
root = tk.Tk()
root.title('사업자 상태 확인')
root.resizable(False, False)
root.geometry('600x75')

# Create Input Widgets
input_frame = Frame(root)

input_text_frame = Frame(
    input_frame,
    background="black",
    bd="0"
)

input_dir_label = Label(
    input_text_frame,
    height="1",
    width="50"
)

# Define Input Button Action


def select_input():
    filetypes = (
        ('엑셀 파일', '*.xlsx'),
        ('모든 파일', '*.*')
    )

    filename = fd.askopenfilename(
        title='사업자등록번호 파일 선택',
        initialdir='./',
        filetypes=filetypes)

    if (filename == "" or filename == None):
        showinfo(
            title='선택한 파일',
            message="파일을 선택하세요"
        )
    else:
        input_dir_label.config(text=filename)


input_file_button = ttk.Button(
    input_frame,
    text='엑셀 파일 선택',
    command=select_input
)

# Render Input Frame
input_frame.pack()
input_text_frame.pack(side="left")
input_dir_label.pack(padx=1, pady=1, side="left")
input_file_button.pack(side="left")


# Create Output Widgets
output_frame = Frame(root)

output_text_frame = Frame(
    output_frame,
    background="black",
    bd=0
)

output_dir_label = Label(
    output_text_frame,
    height="1",
    width="50"
)

# Define Output Button Action


def select_output():
    targetDir = fd.askdirectory(
        title="저장할 위치 선택"
    )

    if (targetDir == "" or targetDir == None):
        showinfo(
            title="지정한 폴더",
            message="폴더를 지정하세요"
        )

    else:
        output_dir_label.config(text=targetDir)


output_dir_button = ttk.Button(
    output_frame,
    text='출력 폴더 선택',
    command=select_output
)

# Render Output Frame
output_frame.pack()
output_text_frame.pack(side="left")
output_dir_label.pack(side="left", padx=1, pady=1)
output_dir_button.pack(side="left")

# Define Check Status Button Action


def checkStatus():
    inputTarget = input_dir_label.cget("text")
    outputTarget = output_dir_label.cget("text")

    if (inputTarget != "" and inputTarget != None and outputTarget != "" and outputTarget != None):
        checkCRNStatus(inputTarget, outputTarget)

        showinfo(
            title="사업자 상태 확인",
            message="조회가 완료되었습니다\n저장위치: " + outputTarget + "/사업자상태.xlsx"
        )

    else:
        showinfo(
            title="항목 누락",
            message="사업자등록번호 엑셀과 출력 파일 저장 폴더를 지정하세요"
        )


# Define Run Widgets
run_check_button = ttk.Button(
    root,
    text='사업자 상태 조회',
    command=checkStatus
)

# Render Run Button
run_check_button.pack(side="right", padx=16)

# Render Main Window
root.configure(background="#ebebeb")
root.mainloop()
