import win32com
from win32com.client import Dispatch
import os

# 指定copy页
page_n = 2
word = win32com.client.Dispatch('Word.Application')

word.Visible = 1 # 后台运行,不显示
word.DisplayAlerts = 0 # 不警告
#	path = # word文件路径
doc_add = word.Documents.Add()
doc_new = word.Documents.Open(r'C:\Users\xx\Desktop\5.doc')
doc = word.Documents.Open(r'C:\Users\xx\Desktop\2.doc', False, False, False)
# word
pages = doc.ActiveWindow.Panes(1).Pages.Count
if page_n > pages:
    print("指定页索引超出已有页面")
else:
    # 123 是word密码 没有则删除或者为''
    objRectangles = doc.ActiveWindow.Panes(1).Pages(page_n)
    # 移动来
    doc.Application.ActiveDocument.Range().GoTo(1, 1, page_n).Select()
    # 记录位置
    start = word.Selection.Start.numerator
    doc.Application.ActiveDocument.Range().GoTo(1, 1, page_n+1).Select()
    # 往左移一下
    word.Selection.MoveLeft()
    if pages==page_n:
        doc.Range().Select()
        word.Selection.MoveRight()
        end = word.Selection.Start.numerator
    else:
        end =  word.Selection.Start.numerator
    doc.Range(start, end).Select()
    word.Selection.Copy()
    doc_new.Application.ActiveDocument.Range().Paste()