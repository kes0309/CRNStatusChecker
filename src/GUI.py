from tkinter import Tk, Label, Frame, Entry, Button, StringVar
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo

from Logic import read_service_key, save_service_key, check_CRN_status

# Create Main Window
root = Tk()
root.title('사업자 상태 확인')
root.resizable(False, False)
root.geometry('800x110')

# Create Service Key Widgets
service_key_label = Label(
    root,
    text="서비스 키: ",
    padx=5,
    pady=3
)

service_key = StringVar()
service_key.set(read_service_key())

service_key_textbox = Entry(
    root,
    textvariable=service_key,
    width=50,
)


def fetch_key():
    key = service_key.get().strip()
    if (key == ""):
        showinfo(
            title="오류",
            message="서비스키를 입력하세요"
        )
        return
    save_service_key(key)


service_key_button = Button(
    root,
    text='서비스키 저장',
    command=fetch_key,
    padx=10
)

service_key_label.grid(row=0, column=0)
service_key_textbox.grid(row=0, column=1)
service_key_button.grid(row=0, column=2)

# Create Input Widgets
input_label = Label(
    root,
    text="사업자등록번호 엑셀: ",
    padx=5,
    pady=3
)

input_text_frame = Frame(
    root,
    bg="black",
    padx=1,
    pady=1
)

input_dir_label = Label(
    input_text_frame,
    bg="white",
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
        initialdir='/',
        filetypes=filetypes)

    input_dir_label.config(text=filename)


input_file_button = Button(
    root,
    text='엑셀 불러오기',
    command=select_input,
    padx=10
)

# Render Input Frame
input_label.grid(row=1, column=0)
input_text_frame.grid(row=1, column=1)
input_dir_label.pack()
input_file_button.grid(row=1, column=2)


# Create Output Widgets
output_label = Label(
    root,
    text="저장할 폴더: ",
    padx=5,
    pady=3
)

output_text_frame = Frame(
    root,
    bg="black",
    padx=1,
    pady=1
)

output_dir_label = Label(
    output_text_frame,
    bg="white",
    width="50"
)

# Define Output Button Action


def select_output():
    targetDir = fd.askdirectory(
        title="저장할 위치 선택"
    )

    output_dir_label.config(text=targetDir)


output_dir_button = Button(
    root,
    text='저장 폴더 선택',
    command=select_output,
    padx=10
)


# Render Output Frame
output_label.grid(row=2, column=0)
output_text_frame.grid(row=2, column=1)
output_dir_label.pack()
output_dir_button.grid(row=2, column=2)

# Define Check Status Button Action


def checkStatus():
    inputTarget = input_dir_label.cget("text")
    outputTarget = output_dir_label.cget("text")
    serviceKey = service_key.get()

    if (serviceKey == ""):
        showinfo(
            title="오류",
            message="서비스키를 입력 후 저장하세요"
        )
    elif (inputTarget == "" or inputTarget == None):
        showinfo(
            title="오류",
            message="사업자등록번호 엑셀을 선택하세요"
        )
    elif (outputTarget == "" or outputTarget == None):
        showinfo(
            title="오류",
            message="출력 파일 저장 위치를 선택하세요"
        )
    else:
        result = check_CRN_status(serviceKey, inputTarget, outputTarget)

        if (result == ""):
            showinfo(
                title="사업자 상태 확인",
                message="조회가 완료되었습니다\n저장된 파일 위치:\n" + outputTarget + "/사업자상태.xlsx"
            )
        else:
            showinfo(
                title="오류",
                message=result
            )


# Define Run Widgets
run_check_button = Button(
    root,
    text='사업자 상태 조회',
    command=checkStatus
)

# Render Run Button
run_check_button.grid(row=3, column=2)

# Render Main Window
root.mainloop()
