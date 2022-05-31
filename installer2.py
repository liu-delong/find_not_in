import os
def get_windows_user_path()->str:
    t=os.popen("reg query \"HKEY_CURRENT_USER\\Environment\" -v path")
    t=t.read()
    t=t.split()
    t=t[3:]
    t=" ".join(t)
    return t

def add_windows_user_path(path_need_add:str)->int:
    org=get_windows_user_path()
    org_list=org.split(";")
    ex=False
    for aorg in org_list:
        if(path_need_add==aorg):
            ex=True
            break
    if ex:
        return os.system("setx path \"{}\"".format(org))
    if org[-1:]==";":
        org+=path_need_add
    else:
        org+=";"+path_need_add
    return os.system("setx path \"{}\"".format(org))
   
    

def set_windows_user_env(key:str,value:str)->int:
    return os.system("setx \"{}\" \"{}\"".format(key,value))

if __name__=="__main__":
    now_path=os.getcwd()
    set_windows_user_env("smj",now_path)
    add_windows_user_path("%smj%")
    print("安装完成！")
    os.system("pause")
