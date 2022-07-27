  #!/usr/bin/env python
# coding: utf-8

# In[38]:


import pandas as pd
import sys
import time
#POPS_2019_202202

# In[3]:
uid=str(input('please insert your UID: '))
file_name=str(input('please insert the input file name : '))

path='C:\\Users\\'+uid+'\\Downloads\\' + file_name
try:
    df = pd.read_csv(path, delimiter=";", sep="\n", encoding="utf-16-le")

except:
    df = pd.read_csv(path, delimiter=";", sep="\n")


def arrondir(ch):

        ch = str(ch)
        
        if ch != "nan":
            try:
                a, b = ch.split(",")
                ch = a + "." + b
                value = float(ch)

            except:
                if ch.isdigit():  # si c'est un entier, round() retourne une erreur
                    value = int(ch)
                else:
                    value = ch
        value = str(round(value, 4))


        return value
# In[76]:


test=True
while test:
    test=False
    debut=str(input("enter THE starting date in the format YYYYMM : "))
    try:
        debut=int(debut)
        if (debut > 204000) or (debut < 200800):
            test=True
            debut = int(input("Date out of bounds, please try again : "))
    except:
        print('bad format')
        test=True


objet=str(input("enter the object type : REAL or BUDGET  "))
while objet not in {"REAL","BUDGET"}:
    objet=str(input("please write \"REAL\" or \"BUDGET\""))


# In[77]:


df=df[df['objet']==objet]


# In[64]:


dates=list(df['date'].unique())
dates.sort()
dates=dates[dates.index(debut)::]


# In[78]:


df


# In[65]:


entities=list(df['EntityName'].unique())
ch='RETL_COLL_14,RETL_RECO_09,TO_OUTST_PCT,COL_FU_OUTST_PCT, EFF_1IMP_PCT'
kpis=ch.split(',')



# In[66]:


df=df.sort_values(by=['EntityName','CodeKPI','date'])


# In[67]:


total=df[df['produit']=='TOTAL']
loans=df[df['produit']=='LOANS']
cards=df[df['produit']=='CARDS']
auto=df[df['produit']=='AUTO']


# In[68]:


#TOTAL is the final excel sheet containing total products
#total is the dataframe extracted from df containing only total products

columns=['EntityName','KPI']+dates
TOTAL=pd.DataFrame(columns=columns)


i=-1
j=0
for pays in entities:
    j+=1
    sys.stdout.write("\r{0}%".format("TOTAL "+str(j*100/len(entities))[:4]))
    sys.stdout.flush()
    time.sleep(0.5)
    a=total[total['EntityName']==pays]
    for kpi in kpis:
        b=a[a['CodeKPI']==kpi]
        TOTAL=TOTAL.append({'EntityName':pays,'KPI':kpi},ignore_index=True)
        i+=1
        
        for date in dates:
            try:
                TOTAL[date].iloc[i]=float(arrondir(b[b['date']==date]['valueLocal'].iloc[0]))
                
            except:
                pass


# In[69]: 


columns=['EntityName','KPI']+dates
LOANS=pd.DataFrame(columns=columns)


i=-1
j=0
for pays in entities:
    j+=1
    sys.stdout.write("\r{0}%".format("LOANS "+str(j*100/len(entities))[:4]))
    sys.stdout.flush()
    time.sleep(0.5)
    a=loans[loans['EntityName']==pays]
    for kpi in kpis:
        b=a[a['CodeKPI']==kpi]
        LOANS=LOANS.append({'EntityName':pays,'KPI':kpi},ignore_index=True)
        i+=1
        for date in dates:
            try:
                LOANS[date].iloc[i]=float(arrondir(b[b['date']==date]['valueLocal'].iloc[0]))
                
            except:
                pass

for pays in entities:
    j+=1
    sys.stdout.write("\r")
# In[70]:


columns=['EntityName','KPI']+dates
CARDS=pd.DataFrame(columns=columns)


i=-1
j=0
for pays in entities:
    j+=1
    sys.stdout.write("\r{0}%".format("CARDS "+str(j*100/len(entities))[:4]))
    sys.stdout.flush()
    time.sleep(0.5)
    a=cards[cards['EntityName']==pays]
    for kpi in kpis:
        b=a[a['CodeKPI']==kpi]
        CARDS=CARDS.append({'EntityName':pays,'KPI':kpi},ignore_index=True)
        i+=1
        for date in dates:
            try:
                CARDS[date].iloc[i]=float(arrondir(b[b['date']==date]['valueLocal'].iloc[0]))
                
            except:
                pass


# In[71]:


columns=['EntityName','KPI']+dates
AUTO=pd.DataFrame(columns=columns)


i=-1
j=0
for pays in entities:
    j+=1
    sys.stdout.write("\r{0}%".format("AUTO "+str(j*100/len(entities))[:4]))
    sys.stdout.flush()
    time.sleep(0.5)
    a=auto[auto['EntityName']==pays]
    for kpi in kpis:
        b=a[a['CodeKPI']==kpi]
        AUTO=AUTO.append({'EntityName':pays,'KPI':kpi},ignore_index=True)
        i+=1
        for date in dates:
            try:
                AUTO[date].iloc[i]=float(arrondir(b[b['date']==date]['valueLocal'].iloc[0]))
                
            except:
                pass


# In[72]:


AUTO.to_excel('pops.xlsx',index=False,sheet_name='AUTO',na_rep="")


# In[74]:


with pd.ExcelWriter (path+".xlsx") as writer:
    TOTAL.to_excel(writer, sheet_name="TOTAL",index=False)
    LOANS.to_excel(writer, sheet_name="LOANS",index=False)
    CARDS.to_excel(writer, sheet_name="CARDS",index=False)
    AUTO.to_excel(writer, sheet_name="AUTO",index=False)


# In[75]:
