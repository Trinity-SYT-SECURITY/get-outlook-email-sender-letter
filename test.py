# import win32com.client  # 讀取郵件模塊
from win32com.client.gencache import EnsureDispatch as DispatchEx
import tkinter as tk
import os

from datetime import datetime, timedelta
from tkinter import filedialog
import re


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
    print("邮件数量：", mails_count)

    received_dt = datetime.now() - timedelta(days=int(dayy))
    received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
    mailss = mails.Restrict("[ReceivedTime] >= '" + received_dt + "'")
    mailss = mails.Restrict("[SenderEmailAddress] = '" + send + "'")
    # mails = mails.Restrict("[UnRead] = 'True'")過濾未讀取郵件
    print(len(mailss))
    outputDir = r"C:\\outputfile" #要改成你想要把郵件中的附件輸出到什麼目錄底下，這裡不能用相對路徑!!
    try:
        for message in list(mailss):
            try:
                #mail = mails.Item(message)
                s = message.SenderName
                b = message.Body

                for attachment in message.Attachments:
                    attachment.SaveAsFile(os.path.join(
                        outputDir, attachment.FileName))
                    print(f"attachment {attachment.FileName} from {s} saved")

                fi = open("output.txt", "a")
                print("attachment {attachment.FileName} from {s} saved \n")
                fi.writelines(b+"\n"+("=="*50))
                print("==" * 50)
                fi.close()
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
        dayy = input("幾天前? (這裡只能輸入數字)")
        if dayy.isdigit():
            break
        else:
            # The input string is not a number
            print('Invalid day,please try again\n')

    #
    """outputDir = input("你想將檔案儲存在哪個目錄下?請輸入目錄路徑!! (EX: C:\\test\) : ").strip()
    print(outputDir)
    user = input("這邊會刪除剛輸入的目錄下所有檔案，如果不要請輸入 n 如果需要輸入 y :").strip()
    # 這樣搞是因為，每次執行程式若把郵件輸出到相同的目錄下，就會堆疊太多紀錄，如果沒需要這功能，可以把它刪掉
    if user == 'y':
        files = os.listdir(outputDir)
    # Iterate through the list of files and use os.remove() to delete each file
        for file in files:
            file_path = os.path.join(outputDir, file)
            os.remove(file_path)

    elif user == 'n':
        print("接下來你所新生成的檔案會在此目錄下，且會包含上一次生成的當案\n")"""
    outlook_mail()
#output = input("輸入要將郵件內容輸出進去的txt檔案名稱 EX: output")
#output = output+'.txt'
#f = open("output.txt", "r+")

# f.truncate()
