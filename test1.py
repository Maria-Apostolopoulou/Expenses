file='DataSet.xlsx'
import pandas as pd
df = pd.read_excel('DataSet.xlsx',sheet_name='October_2018')

print(df.shape)

total_m=0
total_p=0
total_b=0

for i in range(0,len(df)):
    person=df.iloc[i, 3]
    if person=='m':
        total_m=total_m+df.iloc[i, 2]
    elif person=='p':
        total_p=total_p+df.iloc[i,2]
    elif person=='b':
        total_b=total_b+df.iloc[i,2]
    elif person!='m' and person!='m' and person!='m':
        print('Unrecognised Entry', df.iloc[i, 3])

total_all=total_b+total_p+total_m
total_m=total_m+total_b/2
total_p=total_p+total_b/2
total_m_per=(total_m/total_all)*100
total_p_per=(total_p/total_all)*100
print('Marietta = ',total_m)
print('Paul = ', total_p)
print('Both = ',total_b)

import matplotlib.pyplot as plt

# Pie chart, where the slices will be reflecting the %expenditures of each person:
labels = 'Paul', 'Marietta'
sizes = [total_p_per, total_m_per]
colors = ['lightskyblue', 'lightcoral']

explode = (0, 0.1)  # to separate the 2 slices

fig1, ax1 = plt.subplots()

ax1.pie(sizes, explode=explode,labels=labels, colors=colors,autopct='%1.1f%%',
        shadow=True, startangle=90)
ax1.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.

plt.show()

#Analysis of the savings for a 6 month period

df_Jul = pd.read_excel(file,sheet_name='July_2018')
df_Aug = pd.read_excel(file,sheet_name='August_2018')
df_Sep = pd.read_excel(file,sheet_name='September_2018')
df_Oct = pd.read_excel(file,sheet_name='October_2018')
df_Nov = pd.read_excel(file,sheet_name='November_2018')
df_Dec = pd.read_excel(file,sheet_name='December_2018')
Savings=[]
for i in range(0,len(df_Jul)):
    category=df_Jul.iloc[i, 5]
    if category=='savings':
        #Jul_sav=df_Jul.iloc[i, 2]
        Savings=[df_Jul.iloc[i, 2]]
for i in range(0,len(df_Aug)):
    category=df_Aug.iloc[i, 5]
    if category=='savings':
        #Aug_sav=df_Aug.iloc[i, 2]
        Savings.append(df_Aug.iloc[i, 2])
for i in range(0,len(df_Sep)):
    category=df_Sep.iloc[i, 5]
    if category=='savings':
        #Sep_sav=df_Sep.iloc[i, 2]
        Savings.append(df_Sep.iloc[i, 2])
for i in range(0,len(df_Oct)):
    category=df_Oct.iloc[i, 5]
    if category=='savings':
        #Oct_sav=df_Oct.iloc[i, 2]
        Savings.append(df_Oct.iloc[i, 2])
for i in range(0,len(df_Nov)):
    category=df_Nov.iloc[i, 5]
    if category=='savings':
        #Nov_sav=df_Oct.iloc[i, 2]
        Savings.append(df_Nov.iloc[i, 2])
for i in range(0,len(df_Dec)):
    category=df_Dec.iloc[i, 5]
    if category=='savings':
        #Dec_sav=df_Dec.iloc[i, 2]
        Savings.append(df_Dec.iloc[i, 2])

print(Savings)
import statistics
Average_Savings=statistics.mean(Savings)
print(round(Average_Savings,2))
#Plot the progress
Money_in=0
Money=[]
Target_final=2500
Target=[]
for i in range(0,len(Savings)):
    Money_in=Money_in+Savings[i]
    Money.append(round(Money_in,2))
    Target.append(Target_final)

print(Money)
Months=['July','August','September','October','November', 'Decemeber']
x=[1, 2, 3, 4, 5, 6]
import math
numb=math.ceil((Target_final-Money[5])/Average_Savings)
fig, ax = plt.subplots()

plt.scatter(x,Money)
plt.plot(x, Target,'r--')
plt.xlabel('Months in 2018')
plt.ylabel('Cumulative Savings in £')
plt.title('Savings account')
plt.xticks(x,(Months))
ax.text(4,Target[0]-100, 'Target £2500', style='italic', color='red')
ax.text(4,Target[0]-200, 'Months to target:', color='red')
ax.annotate(str(numb),xy=(5.5,Target[0]-200), color='red')

plt.show()

# Breakdown of the personal and professional expenses for both
prof_p=0
prof_m=0
refunds=0
per_m=0
per_p=0

for i in range(0,len(df)):
    person=df.iloc[i, 3]
    type=df.iloc[i, 4]
    if person=='m' and type=='prof':
        prof_m=prof_m+df.iloc[i, 2]
    elif person=='m' and type=='per':
        per_m=per_m+df.iloc[i, 2]
    if person=='p' and type=='prof':
        prof_p = prof_p + df.iloc[i, 2]
    elif person=='p' and type=='per':
        per_p = per_p + df.iloc[i, 2]
    if person=='b' and type=='r':
       refunds=refunds+df.iloc[i,2]
per_m=round(per_m,2)
print("Marietta's breakdown", 'Personal = £',per_m,'Proffesional = £', prof_m)
print("Paul's breakdown", 'Personal = ',per_p,'Proffesional = ', prof_p)
print('Refunds', refunds)

import numpy as np
n_groups = 2
m_expenses=(per_m,prof_m)
p_expenses=(per_p,prof_p)
x_labels=['10%','10%','10%','10%']
fig2, ax2 = plt.subplots()
index = np.arange(n_groups)
bar_width = 0.35
opacity = 0.8
rects1 = plt.bar(index, m_expenses, bar_width,

                 alpha=opacity,
                 color='b',
                 label='Marietta')

rects2 = plt.bar(index + bar_width, p_expenses, bar_width,

                 alpha=opacity,
                 color='g',
                 label='Paul')

plt.xlabel('Type of Expenditure')
plt.ylabel('Cost in £')
plt.title('Expenditures breakdown')
plt.xticks(index+bar_width/2, ('Personal', 'Professional'))



plt.legend()

plt.tight_layout()
plt.show()

# Breakdown of Marietta's expenses
# The categories for personal expenses are [ Health, Dine out, Gifts, Beauty]
#The categories for proffesional expenses are [Commute, Books, Memberships, Conferences, Travel]
# There are 2 groups and 9 subgroups
per_hel=0
per_do=0
per_g=0
per_b=0

prof_com=0
prof_b=0
prof_mem=0
prof_con=0
prof_t=0

for i in range(0,len(df)):
    person=df.iloc[i, 3]
    type=df.iloc[i, 4]
    category=df.iloc[i, 5]
    if person=='m' and type=='prof':
        if category=='com':
            prof_com=prof_com+df.iloc[i, 2]
        elif category=='book':
            prof_b=prof_b+df.iloc[i, 2]
        elif category=='memb':
            prof_mem =prof_mem+df.iloc[i, 2]
        elif category=='conf':
            prof_con=prof_con+df.iloc[i, 2]
        elif category=='travels':
            prof_t=prof_t+df.iloc[i, 2]
    elif person=='m' and type=='per':
        if category=='health':
            per_hel =per_hel+df.iloc[i, 2]
        elif category=='dout':
            per_do=per_do+df.iloc[i, 2]
        elif category == 'gift':
            per_g = per_g + df.iloc[i, 2]
        elif category == 'beauty':
            per_b = per_b + df.iloc[i, 2]


total_per=per_hel+per_do+per_g+per_b
tot_prof=prof_com+prof_b+prof_mem+prof_con+prof_t
#Visualise the data using Nested Pie Charts

import matplotlib.pyplot as plt

group_names=['Professional', 'Personal']
group_size=[tot_prof,total_per]
subgroup_names=['Commute', 'Books', 'Memberships', 'Conferences', 'Airfare', 'Health', 'Eat out','Gift', 'Beauty']
subgroup_size=[prof_com,prof_b,prof_mem,prof_con,prof_t,per_hel, per_do,per_g,per_b]

# Create colors
a, b=[plt.cm.Blues, plt.cm.Reds]

# First Ring (outside)
fig, ax = plt.subplots()
ax.axis('equal')
Pie1, _ = ax.pie(group_size, radius=1.5, labels=group_names, colors=[a(0.6), b(0.6)] )
plt.setp( Pie1, width=0.3, edgecolor='white')

# Second Ring (Inside)
Pie2, _ = ax.pie(subgroup_size, radius=1.5-0.3, labels=subgroup_names, labeldistance=0.7, colors=[a(0.5), a(0.4), a(0.3), a(0.2), a(0.1),  b(0.4), b(0.3), b(0.2), b(0.1)])
plt.setp( Pie2, width=0.4, edgecolor='white')
plt.margins(0,0)

plt.show()
