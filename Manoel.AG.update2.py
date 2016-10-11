#11/10/2016
"""
Created on Mon Oct 03 15:57:12 2016

@author: Bolsistas
"""

import win32com.client as com #Importando pacotes
import os
import random
import pandas as pd

#print("Pacotes carregados!")
#Vissim = com.Dispatch("Vissim.Vissim.800") #Abrindo o Vissim
#print("Vissim aberto")
#Path_of_COM_Basic_Commands_network = os.getcwd() #Formando o caminho de abertura
#EXEMPLOP               = os.path.join(Path_of_COM_Basic_Commands_network, 'C:\Users\Bolsistas.usuario-PC\Leonardo Ribeiro\drive-download-20160919T170903Z\EXEMPLO.inpx')
#flag = False 
#Vissim.LoadNet(EXEMPLOP, flag) #Carregando o arquivo
#print('Arquivo carregado!')
#dados=open('dadospesquisa.txt','w') #Abrindo o arquivo do excel que será editado
#dados.write('Semente  Distancia entre Veiculos  Tempo Minimo de Gap  Delay do Individuo  Velocidade do Individuo\n')
##Simulação
#replicacao=input("Insira o numero de replicacoes...") #Estabelecendo as variáveis inputadas
#ind=input("Insira o numero de individuos por populacao...")
#sodeh=input("Insira a Seed Inicial...")
#y=input("Insira o incremento...")
#delayesp=input("Qual o delay esperado...")
#geracoes=input('Quantas geracoes voce deseja...')
#mg=range(ind)
#ax=range(ind)
#for x in range(ind):
#    mg[x]=round(random.uniform(2.5,5.0),1)
#    ax[x]=round(random.uniform(1,3),1)
#mat=pd.DataFrame({'MinGap':mg,'ax':ax})
#errof=500000000000000 #Estabelecendo variáveis que auxiliarão no cálculo de erros
#errof2=0
dados.write('Semente  Distancia entre Veiculos  Tempo Minimo de Gap  Delay do Individuo  Velocidade do Individuo')
gigasframe=pd.DataFrame({'gerações':[],'individuo':[],'semente':[],'distancia':[],'tmg':[],'delayindidual':[]})
print(gigasframe)

def simulacao(a,b,c,d,e,f,g): #Estabelecendo uma função para a simulação
    #Inputs: a=MinGapTime, b=Semente inicial, c=Ax, d=Incremento da semente, e=Numero de replicações por individuo
    #Outputs: listadel e listavel com valores do delay e velocidade média dessa simulação+Escreve no arquivo csv ou valores de cada semente
    for l in range(replicacao):
        Vissim.Net.PriorityRules[0].ConflictMarkers[0].SetAttValue("MinGapTime", a)
        print("Seed definida:{}".format(b))
        Vissim.Simulation.SetAttValue('RandSeed', b)
        Vissim.Net.DrivingBehaviors[1].SetAttValue('W74ax', c)
        End_of_simulation = 3000 # simulation second [s]
        Vissim.Simulation.SetAttValue('SimPeriod', End_of_simulation)
        print("Simulação Iniciada")
        Vissim.Graphics.CurrentNetworkWindow.SetAttValue("QuickMode",1)            
        Vissim.Simulation.RunContinuous()
        DC_measurement_number = 2
        DC_measurement = Vissim.Net.DataCollectionMeasurements.ItemByKey(DC_measurement_number)
        Speed = DC_measurement.AttValue('Speed(Current,Avg,All)') # Speed of vehicles
        Delay = DC_measurement.AttValue('QueueDelay(Current,Avg,All)') # Length of vehicles
        listamg[l]=a
        listasem[l]=b
        listadv[l]=c
        listavel[l]=Speed
        listadel[l]=Delay
        
        #dados.write('%d;%f;%f;%f;%f\n' % (b, c, a, Delay, Speed))
        b+=d
        
for i in range(ind): #Lista que calculará a média do Delay e da Velocidade por indivíduo
    listavel=range(replicacao)
    listadel=range(replicacao)
    listadv=range(replicacao)
    listasem=range(replicacao)
    listamg=range(replicacao)

    Random_Seed=sodeh
    simulacao(mat['MinGap'][i],sodeh,mat['ax'][i],y,replicacao)
    matvd=pd.DataFrame({'semente':listasem, 'distvel':listadv,'mingap':listamg,'Vel':listavel,'Delay':listadel},columns=['Semente','  Distancia entre Veiculos','  Tempo Minimo de Gap','  Delay do Individuo','  Velocidade do Individuo']) 
    
    velmedia=pd.DataFrame.mean(matvd)['Vel']
    delmedia=pd.DataFrame.mean(matvd)['Delay']
    Random_Seed+=y
    mape=(delmedia-delayesp)/delayesp#a divisão é nova e a terminologia mape tb
    if (abs(errof))>(abs(mape)):
        errof=errom
        dmelhor=delmedia
        vmelhor=velmedia
        indivm=i
    if (abs(mape))>(abs(errof2)):
        errof2=errom
        dpior=delmedia
        vpior=velmedia
        indivp=i    
    dados.write('Velocidade Media; %f\n' % (velmedia)) #Escrevendo os valores médios no arquivo csv
    dados.write('Delay Medio; %f\n' % (delmedia))
    
dados.write('O melhor individuo dessa populacao foi o individuo numero %.0f com velocidade de %.2f e delay de %.2f;\n' % (indivm, vmelhor, dmelhor))
dados.write('O pior individuo dessa populacao foi o individuo numero  %.0f com velocidade de %.2f e delay de %.2f;\n' % (indivp, vpior, dpior))

for r in range(geracoes-1):
    
    #checagem a cada três gerações para predatismo e 
    errof=500000000000000 #Reestabelecendo os erros máximos e mínimos para a filtragem dessa geração
    errof2=0
    dados.write('\n \n')
    dados.write('Geracao %d \n' %(r+2))
    if (r+1)/3 != int((r+1)/3):
        for q in range(ind): #A.G. -> Melhor individuo continua, Pior morre; Outros três cruzam com o melhor; Novo individuo aparece gerado randomicamente
            if q!=indivm:
                if random.random()<.5:
                    mat['ax'][q]=mat['ax'][indivm]
                if random.random()<.5:
                    mat['MinGap'][q]=mat['MinGap'][indivm]
# nesse laço o melhor idividuo sobrevive e cruza com os demais
    else:
        for q in range(ind): #A.G. -> Melhor individuo continua, Pior morre; Outros três cruzam com o melhor; Novo individuo aparece gerado randomicamente
            if q!=indivm:
                if q==indivp:
                    mat['MinGap'][q]=round(random.uniform(2.5,5.0),1)
                    mat['ax'][q]=round(random.uniform(1,3),1)
                else:
                    if random.random()<.5:
                        mat['ax'][q]=mat['ax'][indivm]
                    if random.random()<.5:
                        mat['MinGap'][q]=mat['MinGap'][indivm]
            if random.random()<.2:
                mat['ax'][q]=round(random.uniform(1,3),1)
            if random.random()<.2:
                mat['MinGap'][q]=round(random.uniform(2.5,5.0),1)
               
    for j in range(ind): 
        listavel={}
        listadel={}
        simulacao(mat['MinGap'][j],sodeh,mat['ax'][j],y,replicacao)
        matvd=pd.DataFrame({'Vel':listavel,'Delay':listadel})  
        velmedia=pd.DataFrame.mean(matvd)['Vel']
        delmedia=pd.DataFrame.mean(matvd)['Delay']
        mape=(delmedia-delayesp)/delayesp#a divisão é nova e a terminologia mape tb
        if (abs(errof))>(abs(mape)):
            errof=errom
            dmelhor=delmedia
            vmelhor=velmedia
            indivm=j
        if (abs(mape))>(abs(errof2)):
            errof2=errom
            dpior=delmedia
            vpior=velmedia
            indivp=j    
        dados.write('Velocidade Media; %f\n' % (velmedia))
        dados.write('Delay Medio; %f\n' % (delmedia))
        
    dados.write('O melhor individuo dessa populacao foi o individuo numero %.0f com velocidade de %.2f e delay de %.2f;\n' % (indivm, vmelhor, dmelhor))
    dados.write('O pior individuo dessa populacao foi o individuo numero  %.0f com velocidade de %.2f e delay de %.2f;\n' % (indivp, vpior, dpior))
dados.close()
Vissim = None
print("C'est fini")