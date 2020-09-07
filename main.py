import os
import pandas as pd
import tkinter as tk

def to_one_excel(dir,lists):
    dfs = []        #存放追加内容
    for root_dir, sub_dir, files in os.walk(dir):   # 起始路径，起始路径下的文件夹，起始路径下的文件
        for file in files:                          #文件夹内的所有文件
            if file.endswith('xls') or file.endswith('xlsx'):
                file_name = os.path.join(file)
                # print(file_name)
                rf = list(pd.read_excel(file_name, nrows=1))
                # print(rf)
                indexs = []
                if lists[0] == '$':
                    indexs = rf
                else:
                    #单位编码 姓名* 实发工资
                    for i in lists:
                        # print(i)
                        indexs.append(rf.index(i))
                df = pd.read_excel(file_name,usecols=indexs)
                dfs.append(df)
                text_message.insert('end',file_name+'已处理\n')
                text_message.yview('moveto',1.0)
    df_concated = pd.concat(dfs)
    out_path = os.path.join(dir,'A合并表格.xlsx')
    df_concated.to_excel(out_path, sheet_name='Sheet1', index=None)
    text_message.insert('end', "全部文件处理成功！合并文件为'A合并表格.xlsx'\n")

#绘制GUI
app = tk.Tk()
app.title('AExcel处理工具')

#消息窗口
message_frame = tk.Frame(width=480, height=300,bg='white')  #划分Frame
text_message = tk.Text(message_frame)
message_frame.grid(row=0, column=0, padx=3, pady=6)         #0行0列，边框距离x=3px,y=6px
message_frame.grid_propagate(0)                             #固定面板大小
text_message.grid()

#输入窗口
text_frame = tk.Frame(width=480, height=100)
text_text = tk.Text(text_frame)
text_text.insert('end','$ 输入要合并的列名称，用空格分开，直接点击开始则默认合并所有列')
text_frame.grid(row=1, column=0, padx=3, pady=6)
text_frame.grid_propagate(0)
text_text.grid()

#开始按钮
def start():
    send_msg = text_text.get('0.0',tk.END)                  #获取输入窗口文本内容
    # print(send_msg)
    lists = send_msg.replace('\n','').split(' ')
    # print(lists)
    if lists[0] == '$':
        text_message.insert('end', '开始合并全部列\n')
        to_one_excel(r'./', lists)
    else:
        text_message.insert('end', '开始合并'+send_msg)
        to_one_excel(r'./',lists)
    text_text.delete('0.0',tk.END)                          #清空输入窗口文本内容

send_frame = tk.Frame(width=480, height=30)
button_send = tk.Button(send_frame, text='开始',command=start) #添加按钮并绑定发送功能
send_frame.grid(row=2, column=0, padx=10, sticky='E')       #2行0列,距离边框10px,右对齐
# send_frame.grid_propagate(0)
button_send.grid()

app.mainloop()
