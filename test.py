# import win32com.client  # 讀取郵件模塊
from win32com.client.gencache import EnsureDispatch as DispatchEx
import tkinter as tk
import os

from datetime import datetime, timedelta
from tkinter import filedialog
import re


def write_outputtxt():
    deleted_contents = False

    if os.path.exists("output.txt"):
        # Check if the file is empty
        if os.stat("output.txt").st_size == 0:
            # Write the lines to the file
            with open("output.txt", "a+") as file:
                file.writelines(b+"\n"+("=="*50))
        else:
            # Ask the user if they want to delete the contents of the file
            response = input(
                "Do you want to delete the contents of the file (y/n)? ")

            # If the user wants to delete the contents of the file, do it
            if response.lower() == "y":
                # 如果是第一次輸入 y，則刪除文件內容
                if not deleted_contents:
                    with open("output.txt", "w") as file:
                        file.writelines(b+"\n"+("=="*50))
                    deleted_contents = True
                    # 如果是第二次或之後輸入 y，則不執行任何動作
                else:
                    pass
            elif response.lower() == "n":
                with open("output.txt", "a+") as file:
                    file.writelines(b+"\n"+("=="*50))
# If the file doesn't exist, create it and write the lines to it
    else:
        with open("output.txt", "x") as file:
            file.writelines(b+"\n"+("=="*50))


def outlook_mail():

    # 使用MAPI協議連接Outlook

    # mails_count = len(mails)  # 郵件數量
    #print("郵件數量：", mails_count)
    account = DispatchEx('Outlook.Application').GetNamespace('MAPI')
    # 獲取收件箱所在位置
    inbox = account.GetDefaultFolder(6)  # 數字6代表收件箱
    # 獲取收件箱下所有郵件
    mails = inbox.Items
    mails_count = len(mails)  # 邮件数量
    print("所有的郵件数量：", mails_count)

    received_dt = datetime.now() - timedelta(days=int(dayy))
    received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')

# 限制郵件列表，只保留在給定天數內收到的郵件
    mailstime = mails.Restrict("[ReceivedTime] >= '" + received_dt + "'")

# 再次限制郵件列表，只保留發件人為給定地址的郵件
    mailss = mailstime.Restrict("[SenderEmailAddress] = '" + send + "'")
    # mails = mails.Restrict("[UnRead] = 'True'")過濾未讀取郵件
    print("根據天數及發件人所找到的郵件數量=>"+str(len(mailss))+"\n")

    try:
        for message in list(mailss):
            try:
                #mail = mails.Item(message)
                s = message.SenderName
                global b
                b = message.Body
                # 有附件會執行這for
                for attachment in message.Attachments:
                    attachment.SaveAsFile(os.path.join(
                        outputDir, attachment.FileName))
                    print(f"attachment {attachment.FileName} from {s} saved")

                write_outputtxt()
                #fi = open("output.txt", "a")
                # fi.writelines(b+"\n"+("=="*50))

                print("==" * 60)  # 分隔郵件，不然太多郵件疊加到文件中會太亂
                # fi.close()
            except Exception as e:
                print("error when saving the attachment:" + str(e))
    except Exception as e:
        print("error when processing emails messages:" + str(e))


if __name__ == '__main__':

    pattern = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'

    while True:
        send = input("輸入你想收到的來源發信者的郵件地址:").strip()
        if re.match(pattern, send):
            break
        else:
            # The input string is not a valid email address
            print('Invalid email address,please try again\n')

    while True:
        dayy = input("幾天前? (這裡只能輸入數字) :")
        if dayy.isdigit():
            break
        else:
            # The input string is not a number
            print('Invalid day,please try again\n')

    #
    """outputDir = input("你想將檔案儲存在哪個目錄下?請輸入目錄路徑!! (EX: C:\\test\) : ").strip()
    print(outputDir)"""

    # 記得改成你要把郵件中的附件放到什麼目錄下，不要輸入相對路徑!!!
    outputDir = r"C:\\Users\\user\\Desktop\\郵件分類器\\outputeven"
    user = input("這邊會刪除剛輸入的目錄下所有檔案，如果不要請輸入 n 如果需要輸入 y :").strip()
    # 這樣搞是因為，每次執行程式若把郵件輸出到相同的目錄下，就會堆疊太多紀錄，如果沒需要這功能，可以把它刪掉
    if user == 'y':
        files = os.listdir(outputDir)
    # Iterate through the list of files and use os.remove() to delete each file
        for file in files:
            file_path = os.path.join(outputDir, file)
            os.remove(file_path)

    elif user == 'n':
        print("接下來你所新生成的檔案會在此目錄下，且會包含上一次生成的檔案\n")
    outlook_mail()
#output = input("輸入要將郵件內容輸出進去的txt檔案名稱 EX: output")
#output = output+'.txt'
#f = open("output.txt", "r+")

# f.truncate()
