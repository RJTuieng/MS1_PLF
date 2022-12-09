# -*- coding: utf-8 -*-
"""
Created on Tue Mar 22 11:21:13 2022

@author: rj
"""

import xlsxwriter as xw
import tkinter,sys,itertools, warnings,win32com.client, matplotlib,os, io, math,json
from Bio.SeqUtils import seq3,seq1
from Bio.SeqUtils.ProtParam import ProteinAnalysis
from tkinter.filedialog import askopenfilename
from bioservices.uniprot import UniProt
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from tqdm import tqdm
from Bio import SeqIO
 


def getSeq(protein_id):
    b=u.retrieve(protein_id,frmt='fasta')
    b = b.split('\n')[1:-1]
    protein_seq = ''.join(b)
    return protein_seq

def locatePeptides(seqData,locationPos):

    fragPos=list(seqData['pos'].dropna())
    t=[list(np.zeros(len(sampleList))) for _ in range(len(locationPos))]
    for j in range(len(fragPos)):
        for i in range(len(locationPos)):
            if fragPos[j][0]>=locationPos[i][0] and fragPos[j][0]<=locationPos[i][1] and fragPos[j][1]>=locationPos[i][0] and fragPos[j][1]<=locationPos[i][1]:
                t[i]= list(np.sum([t[i],list(seqData[sampleList].iloc[j])],0))
            elif (fragPos[j][1]>=locationPos[i][0] and fragPos[j][1]<=locationPos[i][1]) or(fragPos[j][0]>=locationPos[i][0] and fragPos[j][0]<=locationPos[i][1]):
                t[i]= list(np.sum([t[i],list(seqData[sampleList].iloc[j])],0))
    return t

def findArrayPos(df,pos):
    shape = np.shape(df)
    start=[int(np.floor((pos[0]-1)/shape[1])),int((pos[0]-1)%shape[1])]
    end = [int(np.floor((pos[1]-1)/shape[1])),int((pos[1]-1)%shape[1])]
    print(f'Start pos at array position [{start[0]},{start[1]}], end pos at array position [{end[0]},{end[1]}]')
    return start, end

def sumNewArray(array,sumPos):
    newDF=np.zeros(len(sumPos))
    for k in range(len(sumPos)):
        newDF[k]=np.average(array[sumPos[k][0]:sumPos[k][1]+1],0)
    return newDF       

def sortfunc(x): #For sorting posList
    y=np.float64(x.split("-")[0].split('[')[1])  
    return y

def domainCleanUp(data):
    testlist=[]
    for x in range(len(data.split('/note='))-1):       
        test1 = data.split('/note=')[x].split(';')[-2].split(' ')[-1]
        if '..' in test1:
            test = test1.split('..')
        else:
            continue
        if ':' in test[0]:
            test[0]=test[0].split(':')[1]
        
        position=[int(test[0]),int(test[1])]
        name= data.split('/note=')[x+1].split('";')[0].strip('"')
        testlist.append([position,name])        
    return testlist

def convertUsefulFasta(protein_data):
    data_list=protein_data.split(' -n ')
    entry_name=data_list[0]
    protein_name=data_list[1]
    protein_seq=data_list[2]
    return entry_name, protein_name, protein_seq
######################################################################################################################

fileLocation = os.getcwd()
# To open dialog to ask for file location 
root=tkinter.Tk()
root.withdraw()
root.wm_attributes('-topmost', 1)
fileName= askopenfilename(parent=root)

u = UniProt()

database ="PythonStuff\\Database\\uniprot_sprot_rat_251022.xlsx"
#for sprot only database #'PythonStuff\\Database\\uniprot_sprot_Human_2022.08.10-09.13.40.25.xlsx'

xls = pd.ExcelFile(database)
df = xls.parse(xls.sheet_names[0])
columns_want = ['Entry Name', 'Protein names','Sequence', 'Domain [CC]', 'Compositional bias',
'Domain [FT]', 'Motif', 'Region']

df_new=pd.DataFrame()
df_new=df[columns_want].apply(lambda x:' -n '.join(x.astype(str)),axis=1)
df_new.index=df.Entry
fasta_sequences=df_new.to_dict()

#Old database (FASTA)
#database='PythonStuff\\Database\\uniprot_sprot_Human_NoIsoform.fasta'
#fasta_sequences = SeqIO.index(database,'fasta',key_function=lambda x: x.split('|')[1])

#Read csv file and arrange headings
book = pd.read_excel(fileName,header=0).fillna(0)
abun_headings=book.columns[book.columns.str.contains('Abundance:')]
norm_abun_headings=book.columns[book.columns.str.contains('(Normalized)')]
sampleList= [x.split(':')[1]+x.split(':')[2] for x in norm_abun_headings]
samplesize=len(sampleList)

#Remove oxidized proteins
#book=book[book['Modifications'].str.contains("1xOxidation"and "\[C")==True]
#book=book[['Oxidation' in str(x) for x in book['Modifications']]]
#'X' in modifications is special: '×'
id_list=book['Master Protein Accessions'].unique()
book['Abundances']=book[abun_headings].apply(lambda x: ','.join(x.astype(str)),axis=1)
book['Abundances (Normalized)']=book[norm_abun_headings].apply(lambda x: ','.join(x.astype(str)),axis=1)
book['Pos']=book['Positions in Master Proteins'].str.split(' ').str[1]
book['Pos_Start']=book['Pos'].str.split('-').str[0].str.split('[').str[1]
book['Pos_End']=book['Pos'].str.split('-').str[1].str.split(']').str[0]
final_head=['Master Protein Accessions','Pos_Start','Pos_End','Abundances (Normalized)','Abundances']
final_book=book[final_head]


#mode either bin or dom
mode='bin'
binSize=20
if mode=='bin':
    outputName= f'Output_{mode}_{binSize}'
else:
    outputName= f'Output_{mode}'

#%%
outputExcel=xw.Workbook(fileName.split('.')[0] + f'_{outputName}.xlsx', {'constant_memory': True})
sheet=outputExcel.add_worksheet()
row=0
for k in tqdm(range(len(id_list))):
    col=3
    protein_id=id_list[k]
    try:
        entry_name, protein_name, protein_seq = convertUsefulFasta(fasta_sequences[protein_id])
    except:
        print(f'{protein_id} not found in database!')
        continue    
    xBook=final_book.loc[final_book['Master Protein Accessions'].str.contains(protein_id,case=False).fillna(False)]
    
    abunDF=np.zeros([len(protein_seq),samplesize])
    
    for h in range(len(xBook)):
        #print(f'{len(x)} repeat for sequence {seqList[0]} \n{x["Modifications"]}')
        #Get peptide location
        
        s=int(xBook['Pos_Start'].iloc[h])-1
        e=int(xBook['Pos_End'].iloc[h])-1 #Convert to zero indexing
        
        abunList=np.zeros(samplesize)
        #Abundances (Normalized)/# PSMs
        #q=xBook['Abundances (Normalized)'].iloc[h].split(',')
        q=xBook['Abundances (Normalized)'].iloc[h].split(',')
        for k in range(samplesize):
            if q[k]=='nan':
                q[k]='0'
        abunList= abunList+np.float64(q)
        abunDF[s:e+1]= abunDF[s:e+1]+abunList  
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", category=RuntimeWarning)
            normFactor=max(sum(abunDF))/sum(abunDF)
        normFactor[normFactor==float('+inf')]=0 #Replace inf into 0
    if mode=='dom':
        
        ###############  Get domain positions  ###############################################
        dList=[]
        for k in range(3,8):
            dList2=domainCleanUp(fasta_sequences[protein_id].split(' -n ')[k])
            dList.append(dList2)
            
        dList = list(itertools.chain(*[x for x in dList if x!=[]]))

        def sort(x):
            return x[0][0]

        dList=sorted(dList,key=sort)

        
        domPos=[x[0] for x in dList] #1 indexing
        domPos=[[s-1,e-1] for s,e in domPos] #0 indexing
        domList=[x[1] for x in dList]
    ################################################################### 
        pos = domPos
    elif mode=='bin':
        binPos=[]
        numBin=len(protein_seq)//binSize
        if numBin==0:
            pos=[0,len(protein_seq)-1]
        else:
            binPos.append([[x,x+binSize-1] for x in range(0,numBin*binSize,binSize)])
            binPos=binPos[0]
            remainder=len(protein_seq)%binSize
            if remainder!=0:
                binPos[-1][1]=binPos[-1][1]+remainder
        pos=binPos
    
    #Write data into excel

    for g in range(samplesize):
        sheet.write(row,col-3,f'{protein_id}')
        sheet.write(row,col-2,f'{sampleList[g]}')
        abun_row_data = sumNewArray(abunDF[:,g], pos)*normFactor[g]
        for m in range(len(pos)):
            try:
                sheet.write(row,col-1,abun_row_data[m])
            except:
                pass
            col+=1        
        
        col=3 #reset column
        row+=1
    
    
    
#    for m in range(len(pos)):
#        if mode=='bin':
#            sheet.write(row,col-1,f'bin:{pos[m]}')
#        if mode=='dom':
#            sheet.write(row,col-1,f'{domList[m]}')
#        col+=1
 #   row+=1

outputExcel.close()   



#%% Stats
print('Now starting stats...')

import pingouin as pg
import pandas as pd
import scipy

    
control_keyword='Control'
treatment_keyword='50Gy'

df=pd.read_excel(fileName.split('.')[0] + f'_{outputName}.xlsx',header=None)


ecm_data=pd.read_excel(f'{fileLocation}//PythonStuff//Database//matrisome_mouse.xls',header=0)

plfAnalysisExcel=xw.Workbook(fileName.split('.')[0] + f'_PLFAnalysis_{control_keyword}VS{treatment_keyword}.xlsx')
sheet1=plfAnalysisExcel.add_worksheet()
headings=['Protein ID','Protein Name']
#for col,headings in enumerate(headings):
sheet1.write(0,0,'Protein ID')
sheet1.write(0,2,'Significant p value?')
sheet1.write(0,3,'Score')
sheet1.write(0,4,'Matrix protein?')
sheet1.write(0,5,'Matrix protein type')
proteinIDList=id_list
row=1


for k in tqdm(range(len(proteinIDList))):
    try:
        e,name,s=convertUsefulFasta(fasta_sequences[proteinIDList[k]])
    except:
        print(f'{proteinIDList[k]} not found in database!')
        continue
    col=7
    sheet1.write(row,0,proteinIDList[k])
    sheet1.write(row,1,name)
    #Use protein_id to find protein name because name is not fully unique
    
    #sheet1.merge_range(row,1,row+5,1,df[df[0]==proteinIDList[k]][1].unique()[0])
    datalist=['pValue','cMean','cStd','tMean','tStd','meanDiff']
    
    for rowlist,head in enumerate(datalist):
        sheet1.write(row+rowlist,col-1,head)
    
    df_specific=df[df[0]==proteinIDList[k]].dropna(axis=1,how='all')
    test=df_specific.melt(id_vars=[0,1])
    #Check if ECM protein
    try:
        ecm_index = ecm_data[ecm_data.UniProt_IDs.str.contains(proteinIDList[k])].index[0]
        sheet1.write(row,4,'Y')
        sheet1.write(row,5,f'{ecm_data.iloc[ecm_index].Category}')
    except:
        sheet1.write(row,4,'N')
        
    #Do anova for each bin
    significant='N'
    score=0
    for x in test['variable'].unique():

        val=test[test['variable']==x]
        cVal=val[val[1].str.contains(control_keyword)]['value']
        tVal=val[val[1].str.contains(treatment_keyword)]['value']
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", category=RuntimeWarning)
            f,p=scipy.stats.f_oneway(cVal,tVal)

        p_corr=p*2
        #p_corr=p*len(test['variable'].unique())  #Bonferroni correction
        if p_corr>1:
            p_corr=1
        
        cMean=np.mean(cVal)
        tMean=np.mean(tVal)
        meanDiff=cMean-tMean
        cStd=np.std(cVal)
        tStd=np.std(tVal)
        
        sheet1.write(row,col,f'{p_corr}')
        sheet1.write(row+1,col,f'{cMean}')
        sheet1.write(row+2,col,f'{cStd}')
        sheet1.write(row+3,col,f'{tMean}')
        sheet1.write(row+4,col,f'{tStd}')
        sheet1.write(row+5,col,f'{meanDiff}')
        col+=1
        if p_corr<=0.05:
            significant='Y'
            score+=1
    sheet1.write(row,2,f'{significant}')
    if len(test['variable'].unique())==0:
        fin_score=0
    else:
        fin_score=score/len(test["variable"].unique())
    sheet1.write_number(row,3,fin_score) #score is weighted by number of bins
    row+=6
    
    
plfAnalysisExcel.close()      






    #%%  Plotting
    #plt.subplot(args, kwargs)
    fig,ax = plt.subplots()
    #ax = fig.add_axes(0,0,1,1)
    xValues=np.arange(len(pos))
    width=0.3
    
    def plotgraph(sample_range):
        newY=[]
        label=sampleList[sample_range[0]].split(', ')[1]
        for k in sample_range:
            newX=sumNewArray(abunDF[:,k], pos) 
            newY.append(newX)
        yValues=np.mean(newY,0)
        sd=np.std(newY,0)
        
        ax.bar(xValues-width,yValues,yerr=sd, label=label,width=width)   
    
    plotgraph(range(5))
    plt.show()

    
#%%    
    
newY=[]
label1='Control'
for k in range(5):
    newX=sumNewArray(abunDF[:,k], pos) 
    newY.append(newX)
   
    #Export data    
    with open(f"{os.getcwd()}\MassSpec\\output\\PDExcel_{label1}_s{k}_output.txt", "w") as txt_file:
        for line in newX: 
            txt_file.write(f'{line}\n')

    print(f'Sample {label1}{k} completed')
    
yValues1=np.mean(newY,0)
sd1=np.std(newY,0)

newY=[]
label2='50Gy'
for k in range(5,10):
    newX=sumNewArray(abunDF[:,k], pos) 
    newY.append(newX)
    
        #Export data    
    with open(f"{os.getcwd()}\MassSpec\\output\\PDExcel_{label2}_s{k}_output.txt", "w") as txt_file:
        for line in newX: 
            txt_file.write(f'{line}\n')
    print(f'Sample {label2}{k} completed')
    
yValues2=np.mean(newY,0)
sd2=np.std(newY,0)

newY=[]
label3='100Gy'
for k in range(10,15):
    newX=sumNewArray(abunDF[:,k], pos) 
    newY.append(newX)
        #Export data    
    with open(f"{os.getcwd()}\MassSpec\\output\\PDExcel_{label3}_s{k}_output.txt", "w") as txt_file:
        for line in newX: 
            txt_file.write(f'{line}\n')
    print(f'Sample {label3}{k} completed')
    
#yValues3=np.mean(newY,0)
#sd3=np.std(newY,0)


ax.bar(xValues-width,yValues1,yerr=sd1, label=label1,width=width)
ax.bar(xValues+width,yValues2,yerr=sd2, label=label2,width=width)
#ax.bar(xValues+0.3,yValues3,yerr=sd3, label='100Gy',width=width)
#%%
with open(f"{os.getcwd()}\MassSpec\\output\\PDExcel_{sampleList[p]}_output.txt", "w") as txt_file:
    for line in x: 
        txt_file.write(f'{line}\n')
print(f'Sample {sampleList[p]} completed')










