# -*- coding: utf-8 -*-
"""
Created on Mon Oct  4 14:20:38 2021

@author: felix.
"""

import requests
import os, sys
from datetime import date, datetime
from PIL import Image, ImageFont, ImageDraw

if __name__ == '__main__':
    
    data_atual = date.today() if len(sys.argv) == 1 else datetime.strptime(sys.argv[1], '%Y%m%d')
    diretorio = '../Mapas' + os.sep + data_atual.strftime('%Y%m%d')
    diretorio_satelite = diretorio + os.sep + 'satelite'
    os.makedirs(diretorio_satelite, exist_ok = True)
    url_base = 'http://satelite.cptec.inpe.br/repositoriogoes/goes16/goes16_web/ams_ret_ch13_baixa/{0}/{1:0>2}/S11635388_{0}{1:0>2}{2:0>2}{3}.jpg'
    ano, mes, dia = data_atual.year, data_atual.month, data_atual.day
    lista_imagens = []
    horarios = ['0000','0100','0200','0300','0400','0500','0600','0700','0800']
    for horario in horarios:
        
        url = url_base.format(ano, mes, dia, horario)
        print(url)
        r = requests.get(url)
        if r.status_code != 200:
            
            print('ERRO: %s'%r.status_code)
            continue
        arq_img = diretorio_satelite + os.sep + 'satelite_{}_{}.png'.format(data_atual, horario)
        with open(arq_img, 'wb') as f: f.write(r.content)
        print('CRIADO: {}'.format(arq_img))
        lista_imagens.append(arq_img)
    
    font = ImageFont.truetype("arial.ttf", 150)
    w1, h1, w2, h2 = 800, 500, 2000, 2000
    imgs = []
    for arq_img in lista_imagens:
        
        img = Image.open(arq_img)
        crop = img.crop((w1, h1, w2, h2))
        draw = ImageDraw.Draw(crop)
        idx = lista_imagens.index(arq_img)
        text = horarios[idx][:2] + ':' + horarios[idx][2:]
        draw.text((100, h2 - h1 - 300), text = text, align = 'center', fill = (0,0,0), font = font)
        if not lista_imagens.index(arq_img): img_saida = crop.copy()
        else: imgs.append(crop)
    arq_saida = diretorio + os.sep + 'satelite00-08.gif'
    img_saida.save(arq_saida, save_all = True, append_images=imgs,optimize=True, duration=900, loop=0)
    img.close()
    print('GIF criado: ' + arq_saida)
