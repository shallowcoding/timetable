#masterpiece2
import pandas as pd
import random
import xlsxwriter as xw
n_o_t=1000
teacher=pd.read_excel('input.xlsx', sheet_name='Sheet1')
t_dict={}
for i in range(len(teacher)):
    key=teacher.loc[i][0]
    if key in t_dict:
        t_dict[key].append(teacher.loc[i][1])
    else:
        t_dict[key]=[]
        t_dict[key].append(teacher.loc[i][1])
teachers=[]
for i in range(len(teacher)):
    teachers.append(teacher.loc[i][0])
subject=pd.read_excel('input.xlsx', sheet_name='Sheet2')
subjects=[]
for i in range(len(subject)):
    subjects.append(subject.loc[i][0])
#print(subject)
ban=pd.read_excel('input.xlsx', sheet_name='Sheet3')
b_num=len(ban.index) #number of classes
b_list=ban.values.tolist() #몇 반이 무슨 과목 듣는지
b_names=ban['반'].values.tolist()
#print(b_names)
b_classes=len(ban.columns)-1
f=open('input.txt', mode='r', encoding='utf-8')
a=[]
lin=-1
for line in f.readlines():
    a.append(line.strip().split(' '))
    lin+=1
t=a[0]
del(a[0])
alp=['MON', 'TUE', 'WED', 'THU', 'FRI']
days={}
for i in range(5):
    days[alp[i]]=int(t[i])

times=[]
n=0
for i in range(5):
    apl=alp[i]
    now=0
    while(now<days[apl]):
        times.append(apl)
        now+=1
    n+=now
must=[]
never=[]
for i in range(lin):
    if (a[i][0] in subjects)==True:
        must.append(a[i])
    if(a[i][0] in teachers)==True:
        never.append(a[i])
oneweek=[[0]*b_classes for i in range(b_num)]
"""
for i in range(b_num):
    for j in range(b_classes):
        for k in range(len(subject)):
            if(subject.loc[k][0]==b_list[i][j+1]):
                dap=k
                break
        oneweek[i][j]=subject.loc[dap][1]
"""
def new(days, b_num, alp, panda):
    class_d={}
    class_d[alp[0]]= [0]*(days[alp[0]])
    class_d[alp[1]]= [0]*(days[alp[1]])
    class_d[alp[2]]= [0]*(days[alp[2]])
    class_d[alp[3]]= [0]*(days[alp[3]])
    class_d[alp[4]]= [0]*(days[alp[4]])
    for j in range(len(alp)): #j 번째 요일
        oneclass=[[] for A in range(days[alp[j]])]
        for i in range(b_num): #i번 반
            t_list=b_list[i][1:]
            for k in range(days[alp[j]]): #j번째 요일의 교시 수
                flag3=0
                for m in range(len(must)):
                    if(must[m][2]==alp[j] and must[m][1]==str(b_list[i][0]) and int(must[m][3])==k+1):
                        tmp=must[m][0]
                        col=0
                        for h in range(j):
                            col+=days[alp[h]]
                        ##print(tmp, i, col+k)
                        panda[i][col+k]=tmp
                        #tt=b_list[i].index(tmp)-1
                        #t_list.remove(tmp)
                        #oneweek[i][tt]-=1
                        oneclass[k].append(sub_to_te(tmp))
                        flag3=1
                        break
                if(flag3==1):
                    continue
                if(len(t_list)==0):
                    print('m')
                    return tmp, oneclass, panda, 0
                tmp=random.choice(t_list)
                tf=check(i, tmp, t_list, oneclass[k], j, k)
                if(tf==1):
                    return tmp, oneclass, panda, 1
                tmp=tf
                col=0
                for h in range(j):
                    col+=days[alp[h]]
                ##print(tmp, i, col+k)
                panda[i][col+k]=tmp
                a=b_list[i].index(tmp)-1
                t_list.remove(tmp)
                oneweek[i][a]-=1
                oneclass[k].append(sub_to_te(tmp))
                ##print('p', panda)
                ##print('o', oneweek)
                #print(oneweek)
    return tmp, oneclass, panda, 0

def sub_to_te(t):
    for h in range(len(list(t_dict.values()))):
        c=list(t_dict.values())[h].count(t)
        if(c>0):
            return list(t_dict.keys())[h]
def check(i, tmp, t_list, oneclass, j, k):
    global b_classes
    c_list=t_list[:]
    nn=0
    for m in range(len(never)):
        if(never[m][1]==alp[j] and never[m][2]==str(k+1)):
            nn=never[m][0]
    while((oneweek[i][b_list[i].index(tmp)-1]<1 or oneclass.count(sub_to_te(tmp))>0) and len(checkn)<b_classes*2 and (tmp!=nn or nn==0)):
        c_list.remove(tmp)
        ##print('c', c_list)
        if(len(c_list)==0):
            return 1
        tmp=random.choice(c_list)
        checkn.append(0)
    ##print('r', tmp)
    ##print('r0', b_list[i])
    ##print(oneweek[i])
    ##print('r1',oneweek[i][b_list[i].index(tmp)-1])
    ##print('r2',oneclass.count(sub_to_te(tmp)))
    if(len(checkn)>=b_classes*2):
        checkn.clear()
        ##print(1)
        return 1
    ##print(len(checkn))
    checkn.clear()
    return tmp
panda_out=[]
checkn=[]
for m in range(n_o_t):
    panda=[[0]*len(times) for i in range(b_num)]
    for i in range(b_num):
        for j in range(b_classes):
            for k in range(len(subject)):
                if(subject.loc[k][0]==b_list[i][j+1]):
                    dap=k
                    break
            oneweek[i][j]=subject.loc[dap][1]
    for j in range(b_num):
        for i in range(len(must)):
            if(must[i][1]==str(b_names[j])):
                oneweek[j][b_list[j].index(must[i][0])-1]-=1
    tmp, oneclass, panda, tfn=new(days, b_num, alp, panda)
    ##print(panda)
    if(tfn==0):
        panda_out.append(panda)
        print('n', end='')
    if(m%100==0):
        print('.', end='')
print(panda_out)
workbook=xw.Workbook('output.xlsx')
worksheet=workbook.add_worksheet()
row=0
col=1
for place in times:
    worksheet.write(row, col, place)
    col+=1
row=1
col=0
for i in range(len(panda_out)):
    for b in b_names:
        worksheet.write(row, col, b)
        row+=1
tmp=0
for i in range(len(panda_out)):
    for j in range(b_num):
        for k in range(len(times)):
            #print(1+j+tmp, 1+k)
            worksheet.write(1+j+tmp, 1+k, panda_out[i][j][k])
    tmp+=b_num
workbook.close()