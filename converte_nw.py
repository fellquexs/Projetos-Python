import os
from datetime import date
from dateutil.relativedelta import relativedelta
from datetime import datetime, timedelta,date
import shutil


year = (datetime.today()+relativedelta(days=15)).strftime('%Y')
month = (datetime.today()+relativedelta(days=15)).strftime('%m')


#Dirs
path = os.path.dirname(os.path.abspath(__file__))
deck_dir = path + os.sep +'decks'+ os.sep +'preliminar nw convertido' + os.sep + year + '-' + month
os.makedirs(deck_dir, exist_ok = True)
deck_dir_copy = path + os.sep +'decks'+ os.sep +'preliminar nw convertido' + os.sep + year + '-' + month +os.sep

#Dirs de Entrada
re = path + os.sep + 'decks' + os.sep + 'ons preliminar' + os.sep + year + '-' + month + os.sep + 'RE.DAT' #alteração do mes

exptDATPMO = path + os.sep + 'decks' + os.sep + 'ons preliminar' + os.sep + year + '-' + month + os.sep + 'EXPT.DAT' #alteração do mes
ccee_aux_EXPTdat = path + os.sep + 'ccee-aux' + os.sep + 'EXPT.DAT'

ccee_aux_TERMdat = path + os.sep + 'ccee-aux' + os.sep + 'TERM.DAT'
termDATPMO = path + os.sep + 'decks' + os.sep + 'ons preliminar' + os.sep + year + '-' + month + os.sep + 'TERM.DAT' #alteração do mes

arq_manauara = path + os.sep + 'ccee-aux' + os.sep + 'MANAUARA.DAT'

DecksNWPrePMO = path + os.sep + 'decks' + os.sep + 'ons preliminar' + os.sep + year + '-' + month + os.sep #origem para mover arq exeto o que estao na listoff



#Dirs de Saída
save_re = deck_dir + os.sep + 'RE.DAT'
save_EXPTdat = deck_dir + os.sep + 'EXPT.DAT'
save_TERMdat = deck_dir + os.sep + 'TERM.DAT'

#Listas
lista_REdat = [' 10']
lista_EXPTdat = [' 201 GTMIN',' 322 GTMIN',' 323 GTMIN',' 324 GTMIN',' 325 GTMIN',' 326 GTMIN',' 327 GTMIN',' 205 GTMIN',' 140 GTMIN',' 328 GTMIN',' 329 GTMIN',' 330 GTMIN']
lista_TERMdat = [' 201',' 140']
lista_TERMdat2 = [' 201',' 140',' 205']
listoff = ['RE.DAT','TERM.DAT','EXPT.DAT'] #para copiar os arq exeto esses pois já estão convertidos

#----------------------------------------------Conversão do RE.DAT--------------------------------------------------------------------------------------------------------------------

bloco_re = open(re)
bloco_re = bloco_re.readlines()

for restricao_REdat in lista_REdat:
    for linREdat in bloco_re[:]:
        if linREdat.startswith(restricao_REdat):
            bloco_re.remove(linREdat)
conv_re = ''.join(bloco_re)
print('-----------------------------------------------Bloco preliminar RE.DAT convertido com sucesso.----------------------------------------------------------------------------')
print(conv_re)

with open(save_re,'w') as f:
    f.write(conv_re)

#####################################################################################################################################################################################
#----------------------------------------------Conversão do EXPT.DAT--------------------------------------------------------------------------------------------------------------------


bloco_EXPTdat = open(ccee_aux_EXPTdat)
bloco_EXPTdat = bloco_EXPTdat.readlines()

file2 = []
for restricao_EXPTdat in lista_EXPTdat:
     for linEXPTdat in bloco_EXPTdat:
         if linEXPTdat.startswith(restricao_EXPTdat):
             recortesGTMIN = ''.join(linEXPTdat)
             file2.append(recortesGTMIN)
recortesGTMIN = ''.join(file2)

bloco_EXPTdat_pmo = open(exptDATPMO) #Bloco EXPTDAT PMO
bloco_EXPTdat_pmo = bloco_EXPTdat_pmo.readlines()

for restricao_EXPTdat2 in lista_EXPTdat:
    for linEXPTdat2 in bloco_EXPTdat_pmo[:]:
        if linEXPTdat2.startswith(restricao_EXPTdat2):
            bloco_EXPTdat_pmo.remove(linEXPTdat2)
recortesONS = ''.join(bloco_EXPTdat_pmo)

conv_EXPTdat = recortesONS + recortesGTMIN

print('-----------------------------------------------Bloco preliminar EXPT.DAT convertido com sucesso.----------------------------------------------------------------------------')
print(conv_EXPTdat)

with open(save_EXPTdat,'w') as file:
    file.write(conv_EXPTdat)

#####################################################################################################################################################################################
#----------------------------------------------Conversão do TERM.DAT--------------------------------------------------------------------------------------------------------------------

bloco_TERMdat = open(ccee_aux_TERMdat)
bloco_TERMdat = bloco_TERMdat.readlines()

file4 = []
for termicas in lista_TERMdat:
     for lint in bloco_TERMdat:
         if lint.startswith(termicas):
             recortesRESTR201140 = ''.join(lint)
             file4.append(recortesRESTR201140)
recortesRESTR201140 = ''.join(file4)

bloco_TERMdat_pmo = open(termDATPMO)
bloco_TERMdat_pmo = bloco_TERMdat_pmo.readlines()


for termicas1 in lista_TERMdat2:
    for lint1 in bloco_TERMdat_pmo[:]:  
        if lint1.startswith(termicas1):
            bloco_TERMdat_pmo.remove(lint1) 
recortes_pmo = ''.join(bloco_TERMdat_pmo)
recortesPARCIAL = recortes_pmo + recortesRESTR201140

with open(save_TERMdat,'w') as f:
    f.write(recortesPARCIAL)

with open(arq_manauara, 'r') as file:
    manauara = file.readlines()

with open(save_TERMdat, 'r') as file:
    recortesPARCIAL = file.readlines()

conv_TERMdat = recortesPARCIAL + manauara
conv_TERMdat = ''.join(conv_TERMdat)

print(conv_TERMdat)

with open(save_TERMdat,'w') as file:
    file.write(conv_TERMdat)

arquivos = os.listdir(DecksNWPrePMO)
aux_arquivos = arquivos

for restricoes in listoff:
    if  restricoes in arquivos[:]:
        aux_arquivos.remove(restricoes)


  
allfiles = aux_arquivos
  
for f in allfiles:
    print(f'movendo files {f}\n================')
    shutil.copy(DecksNWPrePMO + f, deck_dir_copy + f)

