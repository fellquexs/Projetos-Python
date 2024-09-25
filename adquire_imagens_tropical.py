# -*- coding: utf-8 -*-
"""
Created on Mon Jul  6 12:36:40 2020

@author: rene.yoshida
"""

import requests
import os
from datetime import date, timedelta
from PIL import Image
from glob import glob
from pathlib import Path

def adquire_imagens_cfs_semanal(data_atual, diretorio_cfs_semanal):
    
    os.makedirs(diretorio_cfs_semanal, exist_ok = True)
    url = 'https://www.tropicaltidbits.com/analysis/models/cfs-avg/{0}00/cfs-avg_apcpna_samer_{1}.png'
    arq_cfs_semanal = diretorio_cfs_semanal + os.sep + 'cfs_semanal{}.png'
    for n in range(1,7):
        
        r = requests.get(url.format(data_atual.strftime('%Y%m%d'), n))
        with open(arq_cfs_semanal.format(n), 'wb') as f: f.write(r.content)
    print('CFS 00z semanal capturado.')
    
    return

def adquire_imagens_cfs_semanal_06z(data_atual, diretorio_cfs_semanal_06z):
    
    os.makedirs(diretorio_cfs_semanal_06z, exist_ok = True)
    url = 'https://www.tropicaltidbits.com/analysis/models/cfs-avg/{0}06/cfs-avg_apcpna_samer_{1}.png'
    arq_cfs_semanal_06z = diretorio_cfs_semanal_06z + os.sep + 'cfs_semanal{}.png'
    for n in range(1,7):
        
        r = requests.get(url.format(data_atual.strftime('%Y%m%d'), n))
        with open(arq_cfs_semanal_06z.format(n), 'wb') as f: f.write(r.content)
    print('CFS 06z semanal capturado.')
    
    return
    
def adquire_imagens_cfs_semanal_12z(data_atual, diretorio_cfs_semanal_12z):
    
    os.makedirs(diretorio_cfs_semanal_12z, exist_ok = True)
    url = 'https://www.tropicaltidbits.com/analysis/models/cfs-avg/{0}12/cfs-avg_apcpna_samer_{1}.png'
    arq_cfs_semanal_12z = diretorio_cfs_semanal_12z + os.sep + 'cfs_semanal{}.png'
    for n in range(1,7):
        
        r = requests.get(url.format(data_atual.strftime('%Y%m%d'), n))
        with open(arq_cfs_semanal_12z.format(n), 'wb') as f: f.write(r.content)
    print('CFS 12z semanal capturado.')
    
    return
 
def adquire_imagens_cfs_semanal_18z(data_atual, diretorio_cfs_semanal_18z):
    
    os.makedirs(diretorio_cfs_semanal_18z, exist_ok = True)
    url = 'https://www.tropicaltidbits.com/analysis/models/cfs-avg/{0}18/cfs-avg_apcpna_samer_{1}.png'
    arq_cfs_semanal_18z = diretorio_cfs_semanal_18z + os.sep + 'cfs_semanal{}.png'
    for n in range(1,7):
        
        r = requests.get(url.format(data_atual.strftime('%Y%m%d'), n))
        with open(arq_cfs_semanal_18z.format(n), 'wb') as f: f.write(r.content)
    print('CFS 18z semanal capturado.')
    
    return 
def adquire_imagens_cfs_mensal(data_atual, diretorio_cfs_mensal):
    
    os.makedirs(diretorio_cfs_mensal, exist_ok = True)
    url = 'https://www.tropicaltidbits.com/analysis/models/cfs-mon/{0}00/cfs-mon_01_apcpna_month_samer_{1}.png'
    arq_cfs_mensal = diretorio_cfs_mensal + os.sep + 'cfs_mensal{}.png'
    for n in range(1,5):
        
        r = requests.get(url.format(data_atual.strftime('%Y%m%d'), n))
        with open(arq_cfs_mensal.format(n), 'wb') as f: f.write(r.content)
    print('CFS mensal capturado.')

    return

def adquire_imagens_gefs_semanal(data_atual, diretorio_gefs):
    
    os.makedirs(diretorio_gefs, exist_ok = True)
    url = 'https://www.tropicaltidbits.com/analysis/models/gfs-ens/{0}00/gfs-ens_apcpna_samer_{1}.png'
    arq_gefs_semanal = diretorio_gefs + os.sep + 'gefs_semanal{}.png'
    for n in range(5):
        
        r = requests.get(url.format(data_atual.strftime('%Y%m%d'), n * 7 + 1))
        with open(arq_gefs_semanal.format(n), 'wb') as f: f.write(r.content)
    print('GEFS extendido semanal capturado.')
    
    return

def concatena_imagens(diretorio_mapa, w1, h1, w2, h2):
    
    imagens = glob(diretorio_mapa + '/*.png')
    splits = diretorio_mapa.split(os.sep)
    mapa = splits[-1]
    diretorio_ant = os.sep.join(splits[:-1])

#    w1, h1, w2, h2 = 40, 40, 760, 760  
    total_width = (w2 - w1) * len(imagens) 
    max_height = h2 - h1
    comp = Image.new('RGB', (total_width, max_height))
    x_offset = 0
    for imagem in imagens:
        
        img = Image.open(imagem)
        crop = img.crop((w1, h1, w2, h2))
        comp.paste(crop, (x_offset,0))
        x_offset += (w2 - w1)
    comp.save(diretorio_ant + os.sep + '{}_agrupado.png'.format(mapa))
    print('Imagem concatenada criada!')

if __name__ == '__main__':

    # Cria diretorio da data atual
    n_dias = int(input('Numero de dias em relacao a hoje (n): '))
    data_atual = date.today() - timedelta(n_dias)
    diretorio_saida = '../Mapas' + os.sep + data_atual.strftime('%Y%m%d')
    os.makedirs(diretorio_saida, exist_ok = True)
    #=========CFS===========
    # Semanal 00z 
    diretorio_cfs_semanal = diretorio_saida + os.sep + 'cfs_semanal'
    print(os.sep)
    adquire_imagens_cfs_semanal(data_atual, diretorio_cfs_semanal)
    file_size = os.path.getsize(diretorio_cfs_semanal + os.sep + 'cfs_semanal1.png')
    #file_size = os.path.getsize(r'C:\Users\DanielleAparecidadaM\2W ENERGIA S.A\Ferramentas - Ferramentas\1_MIDDLE_2W\Mapas\20220809\cfs_semanal\cfs_semanal1.png')
    print(file_size)
    if file_size > 10240:
        concatena_imagens(diretorio_cfs_semanal, w1 = 40, h1 = 40, w2 = 760, h2 = 760)
    #=========CFS===========
#    # Semanal 06z
    elif file_size > 10240 :
        diretorio_cfs_semanal_06z = diretorio_saida + os.sep + 'cfs_semanal_06z'
        adquire_imagens_cfs_semanal_06z(data_atual, diretorio_cfs_semanal_06z)
        file_size = os.path.getsize(diretorio_cfs_semanal_06z + os.sep + 'cfs_semanal1.png')
        print(file_size)
        if file_size > 10240:
            concatena_imagens(diretorio_cfs_semanal_06z, w1 = 40, h1 = 40, w2 = 760, h2 = 760)
#    #=========CFS===========
#    # Semanal 12z
    elif file_size > 10240:
        diretorio_cfs_semanal_12z = diretorio_saida + os.sep + 'cfs_semanal_12z'
        adquire_imagens_cfs_semanal_12z(data_atual, diretorio_cfs_semanal_12z)
        file_size = os.path.getsize(diretorio_cfs_semanal_12z + os.sep + 'cfs_semanal1.png')
        print(file_size)
        if file_size > 10240:
            concatena_imagens(diretorio_cfs_semanal_12z, w1 = 40, h1 = 40, w2 = 760, h2 = 760)
#    #=========CFS===========
#    # Semanal 18z
    else:
        diretorio_cfs_semanal_18z = diretorio_saida + os.sep + 'cfs_semanal_18z'
        adquire_imagens_cfs_semanal_18z(data_atual, diretorio_cfs_semanal_18z)
        file_size = os.path.getsize(diretorio_cfs_semanal_18z + os.sep + 'cfs_semanal1.png')
        print(file_size)
        if file_size > 10240:
            concatena_imagens(diretorio_cfs_semanal_18z, w1 = 40, h1 = 40, w2 = 760, h2 = 760)
    # Mensal
    diretorio_cfs_mensal = diretorio_saida + os.sep + 'cfs_mensal'
    adquire_imagens_cfs_mensal(data_atual, diretorio_cfs_mensal)
    concatena_imagens(diretorio_cfs_mensal, w1 = 40, h1 = 70, w2 = 760, h2 = 660)
    #=========GEFS==========
    if n_dias > 0:
        
        diretorio_gefs = diretorio_saida + os.sep + 'gefs'
        adquire_imagens_gefs_semanal(data_atual, diretorio_gefs)
        concatena_imagens(diretorio_gefs, w1 = 40, h1 = 40, w2 = 760, h2 = 760)
    
    print('Processo terminado!')