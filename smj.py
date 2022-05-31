import os
import openpyxl
def ismatch(person,file_name):
    for m in person:
        if file_name.find(m)!=-1:
            return True
    return False

if __name__=="__main__":
    now_path=os.getcwd()
    
    root_path=os.environ.get('smj')

    person_messages=[]
    
    wb=None
    try:
        wb = openpyxl.load_workbook(os.path.join(now_path,"名单.xlsx"))
        print("当前目录下有《名单.xlsx》，使用当前目录下的名单")
    except:
        try:
            wb = openpyxl.load_workbook(os.path.join(root_path,"名单.xlsx"))
            print("当前目录下没有《名单.xlsx》，使用安装目录({})下的《名单.xlsx》".format(root_path))
        except:
            print("在当前目录找不到名单.xlsx，安装目录：({})也没有《名单.xlsx》".format(root_path))    
            exit(0)
    ws=wb.worksheets[0]     
    for row in ws.rows:
        person=[]
        for col in row:
            value=col.value
            if value:
                person.append(str(value))
        if len(person)>0:
            person_messages.append(person)
    wb.close()
    person_submit_count=[[] for i in range(len(person_messages))]
    folder=os.listdir(".")
    for file in folder:
        for i in range(len(person_messages)):
            if(ismatch(person_messages[i],file)):
                person_submit_count[i].append(file)
    print("未交名单：")
    for i in range(len(person_messages)):
        if(len(person_submit_count[i])==0):
            print("_".join(person_messages[i]))
    
    print("重复提交名单：")
    for i in range(len(person_messages)):
        if(len(person_submit_count[i])>1):
            print("_".join(person_messages[i])+" "+str(person_submit_count[i]))

    os.system("pause")
