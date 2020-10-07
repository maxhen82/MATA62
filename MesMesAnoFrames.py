import pandas as pd
import matplotlib
import numpy
import statistics

def contagemVoo(lista1,lista2,tipo):
    referencia = lista1
    comparador = lista2
    conta = tipo
    listaContagem = []
    listaReal = []
    listaPerc = []
    listaCancel = []
    
    
    for i in range(len(referencia)):
        contavoo=0
        contareal=0
        for j in range(len(comparador)):
            if(referencia[i]==comparador[j]):
                contavoo+=1
                if ((sit_Voo[j] == "Realizado" or sit_Voo[j] =="REALIZADO") and sit_Voo[j] != ' '):
                    contareal+=1
        listaContagem.append(contavoo)
        listaReal.append(contareal)
        percentual=(round(contareal/contavoo*100,2))
        listaPerc.append(percentual)
        listaCancel.append(round(100-percentual,2))
        #print(referencia[i],'=', contavoo,'->', contareal)
    if (tipo == 'total'):
        #print("lista de contagem: ", listaContagem)
        return (listaContagem)
    if (tipo == 'real'):
        #print("lista Realizados: ",listaReal)
        return (listaReal)
    if (tipo == 'perc'):
        #print("lista % Realizados: ",listaPerc)
        return (listaPerc)
    if (tipo == 'canc'):
        return (listaCancel)


mes=0
mensal=0
anual = []
col_Emp = []
sit_Voo = []
aero_Origem = []
aero_Destino = []
just_Voo = []

for ano in range(0,5,1):
    print('Processando ano ', ano+2015,'...', '\n')
    for mes in range(12):
        print('Processando mês ', mes+1,'...', '\n')
        if (ano==0):
            planilha_1 = pd.read_excel(r"G:\EngSoft\vra2015completo.xlsx",mes)
        elif (ano==1):
            planilha_1 = pd.read_excel(r"G:\EngSoft\vra2016completo.xlsx",mes)
        elif (ano==2):
            planilha_1 = pd.read_excel(r"G:\EngSoft\vra2017completo.xlsx",mes)
        elif (ano==3):
            planilha_1 = pd.read_excel(r"G:\EngSoft\vra2018completo.xlsx",mes)
        else: 
            planilha_1 = pd.read_excel(r"G:\EngSoft\vra2019completo.xlsx",mes)
        

        #Processamento dados
        col_Emp_mes = planilha_1['Sigla  da Empresa'] #empresas presentes
        sit_Voo_mes = planilha_1['Situação'] #status do voo 
        aero_Origem_mes = planilha_1['Aeroporto Origem'] #aeroportos de partida
        aero_Destino_mes = planilha_1['Aeroporto Destino'] #aeroportos de chegada
        
        

        col_Emp.extend(col_Emp_mes)
        sit_Voo.extend(sit_Voo_mes) 
        aero_Origem.extend(aero_Origem_mes) 
        aero_Destino.extend(aero_Destino_mes)
        
        # Voo mensal
        aba_mes = str(mensal)+'_' + str(mes+1)+'-' + str(ano+15)
        realizado_mes=0
        justificado_mes=0
        vooGeral_mes = len(sit_Voo_mes)
        for voo in range(vooGeral_mes):
            if ((sit_Voo_mes[voo] == "Realizado" or sit_Voo_mes[voo] =="REALIZADO") and sit_Voo_mes[voo] != ' '):
                realizado_mes+=1
            #if(just_Voo_mes[voo]!=' '):
                    #justificado+=1
        percGeral_mes = round(realizado_mes/vooGeral_mes*100,2)
        percGeralCan_mes = round(100-percGeral_mes,2)
        



        frameGeral_mes = pd.DataFrame({'Programado' : [vooGeral_mes],
                                  'Realizados' : [realizado_mes],
                                  '% Realizados' : [percGeral_mes],
                                  '% Cancelados' : [percGeralCan_mes],
                                   
                                  })

        print('\n' ,'frameGeral_mes', '\n', frameGeral_mes, '\n')
        
        
        if mensal==0:
            with pd.ExcelWriter('InfraeroMensal.xlsx') as writer:
                frameGeral_mes.to_excel(writer,sheet_name= aba_mes, index=False)
        else:
            with pd.ExcelWriter('InfraeroMensal.xlsx', mode='a') as writer:
                frameGeral_mes.to_excel(writer,sheet_name= aba_mes, index=False)
        mensal+=1
        



        
        print('Mês ', mes+1, 'ano ', ano+2015, ' processado.', '\n')
   


    
    #final do for mes***************************
    print('Processando ano ', ano+2015,' finalizado!', '\n')

    #Processamento Geral
    #print("Voo Mes",'\n')
    realizado=0
    justificado=0
    vooGeral = len(sit_Voo)
    for voo in range(vooGeral):
        if ((sit_Voo[voo] == "Realizado" or sit_Voo[voo] =="REALIZADO") and sit_Voo[voo] != ' '):
            realizado+=1
        #if(just_Voo[voo]!=' '):
                #justificado+=1
    percGeral = round(realizado/vooGeral*100,2)
    percGeralCan = round(100-percGeral,2)
    #percJustificado = round(justificado/realizado*100,2)
    #print(justificado, ':',realizado)
    #print("Total de Voos", '\n', "Programado : Realizados : %Realizados : %Cancelados : % Justificados", '\n')
    #print(vooGeral,' : ',  realizado,' : ', percGeral,' : ', percGeralCan) 



    frameGeral = pd.DataFrame({'Programado' : [vooGeral],
                              'Realizados' : [realizado],
                              '% Realizados' : [percGeral],
                              '% Cancelados' : [percGeralCan],
                              #'% Jutificados' : [percJustificado]  
                              })

    print('\n' ,'frameGeral', '\n', frameGeral)








    print('\n', "********************************************************************",'\n')        

    #Processamento Empresas
    print("Empresas",'\n')
    Emp_sem_repeticoes = list(set(col_Emp))
    vooEmp=[]
    vooReal = []
    vooPercReal= []
    vooPerCan=[]
    vooEmp = contagemVoo(Emp_sem_repeticoes, col_Emp, 'total')
    #print("Total de voos:", '\n', vooEmp,'\n')
    vooReal = contagemVoo(Emp_sem_repeticoes,col_Emp, 'real')
    #print("Voos Realizados:", '\n', vooReal,'\n')
    vooPercReal= contagemVoo(Emp_sem_repeticoes, col_Emp, 'perc')
    #print("Percentual de Realizados:", '\n', vooPercReal,'\n')
    vooPerCan = contagemVoo(Emp_sem_repeticoes, col_Emp, 'canc')
    



    

    frameEmpresas = pd.DataFrame({'Empresa' : Emp_sem_repeticoes,
                                  'Programados' : vooEmp,
                                  'Realizados' : vooReal,
                                  '% Realizados' : vooPercReal,
                                  '% Cancelados' : vooPerCan,
                                  })
    print('\n' ,'frameEmpresas','\n',frameEmpresas)

    a=[]
    b=[]
    c=[]
    d=[]
    e=[]
    tam=len(Emp_sem_repeticoes)
    mediana = statistics.median_high(vooPercReal)
    media = round(statistics.mean(vooEmp),0)
    for corte in range(tam):
        if vooReal[corte]>mediana and vooEmp[corte] > media:
            a.append(Emp_sem_repeticoes[corte])
            b.append(vooEmp[corte])
            c.append(vooReal[corte])
            d.append(vooPercReal[corte])
            e.append(vooPerCan[corte])
            
            
    a.append('Outros')
    b.append(sum(vooEmp)-sum(b))
    c.append(sum(vooReal)-sum(c))
    dd=c[len(c)-1]/b[len(b)-1]
    d.append(round(dd*100,2))
    e.append(100-round(dd*100,2))

    
    
    frameCorteEmpresa = pd.DataFrame({'Empresa' : a,
                                        'Programados' : b,
                                        'Realizados' : c,
                                        '% Realizados' : d,
                                        '% Cancelados' : e
                                        })
    print('\n' ,'Corte media voos:', media , '\n' , 'corte mediana realizados:', mediana ,'\n','frameCorteEmpresa', '\n' , frameCorteEmpresa, '\n', '\n')



    

    print('\n', "********************************************************************",'\n')

    #Processamento Aeroportos de Origem
    print("Aeroporto Partida",'\n')
    Origem_sem_repeticoes = list(set(aero_Origem))
    #print(Origem_sem_repeticoes,'\n')


    vooOrigem = [] # (Origem_sem_repeticoes,aero_Origem, 'total_Origem')
    origemReal = []
    vooOriReal = []
    vooOriCan = []

    vooOrigem = contagemVoo(Origem_sem_repeticoes,aero_Origem, 'total')
    #print(vooOrigem, '\n')
    origemReal = contagemVoo(Origem_sem_repeticoes,aero_Origem, 'real')
    #print(origemReal, '\n')
    vooOriReal= contagemVoo(Origem_sem_repeticoes,aero_Origem, 'perc')
    #print(vooOriReal, '\n')
    vooOriCan= contagemVoo(Origem_sem_repeticoes,aero_Origem, 'canc')
    
    
    
    frameAeroOrigem = pd.DataFrame({'Aeroporto' : Origem_sem_repeticoes,
                                    'Programados' : vooOrigem,
                                    'Realizados' : origemReal,
                                    '% Realizados' : vooOriReal,
                                    '% Cancelados' : vooOriCan,
                                    'Mediana': mediana
                                    })

    
    print('\n' ,'frameAeroOrigem','\n',frameAeroOrigem)
    
    a=[]
    b=[]
    c=[]
    d=[]
    e=[]

    tam=len(Origem_sem_repeticoes)
    mediana = statistics.median_high(vooOriReal)
    media = round(statistics.mean(vooOrigem),0)
    for corte in range(tam):
        if (vooOriReal[corte]>mediana and vooOrigem[corte]>media):
            a.append(Origem_sem_repeticoes[corte])
            b.append(vooOrigem[corte])
            c.append(origemReal[corte])
            d.append(vooOriReal[corte])
            e.append(vooOriCan[corte])

    a.append('Outros')
    b.append(sum(vooOrigem)-sum(b))
    c.append(sum(origemReal)-sum(c))
    dd=c[len(c)-1]/b[len(b)-1]
    d.append(round(dd*100,2))
    e.append(100-(round(dd*100,2)))

    
    frameCorteAeroOrigem = pd.DataFrame({'Aeroporto' : a,
                                    'Programados' : b,
                                    'Realizados' : c,
                                    '% Realizados' : d,
                                    '% Cancelados' : e
                                    })
    
    print('\n' ,'Corte media voos:', media , '\n' , 'corte mediana realizados:',
          '\n' , 'frameCorteAeroOrigem', '\n' , frameCorteAeroOrigem, '\n')


    print('\n', "********************************************************************",'\n')

    #Processamento Aeroportos de Chegada
    print("Aeroportos de Chegada",'\n')
    Destino_sem_repeticoes = list(set(aero_Destino))
    #print(Destino_sem_repeticoes)
    vooDestino= []
    destinoReal= []
    vooDesReal= []
    vooDesCanc= []

    vooDestino = contagemVoo(Destino_sem_repeticoes,aero_Destino, 'total')
    #print(vooOrigem, '\n')
    destinoReal = contagemVoo(Destino_sem_repeticoes,aero_Destino, 'real')
    #print(origemReal, '\n')
    vooDesReal= contagemVoo(Destino_sem_repeticoes,aero_Destino, 'perc')
    #print(vooOriReal, '\n')
    vooDesCanc= contagemVoo(Destino_sem_repeticoes,aero_Destino, 'canc')
    
    
    frameAeroDestino = pd.DataFrame({'Destino' : Destino_sem_repeticoes,
                                    'Programados' : vooDestino,
                                    'Realizados' : destinoReal,
                                    '% Realizados' : vooDesReal,
                                    '% Cancelados' : vooDesCanc,
                                    
                                    })
    print('\n','frameAeroDestino' , '\n' ,frameAeroDestino)
    
    a=[]
    b=[]
    c=[]
    d=[]
    e=[]
    tam=len(vooDesReal)
    mediana = statistics.median_high(vooDesReal)
    media = round(statistics.mean(vooDestino),0)
    for corte in range(tam):
        if (vooDesReal[corte]>mediana and vooDestino[corte]>media):
            
            a.append(Destino_sem_repeticoes[corte])
             
            b.append(vooDestino[corte])
            
            c.append(destinoReal[corte])
            
            d.append(vooDesReal[corte])
            
            e.append(vooDesCanc[corte])
            

    a.append('Outros')
    b.append(sum(vooDestino)-sum(b))
    c.append(sum(destinoReal)-sum(c))
    dd=c[len(c)-1]/b[len(b)-1]
    d.append(round(dd*100,2))
    e.append(100-(round(dd*100,2)))
    
            


    frameCorteAeroDestino = pd.DataFrame({'Destino' : a,
                                    'Programados' : b,
                                    'Realizados' : c,
                                    '% Realizados' : d,
                                    '% Cancelados' : e
                                    })
    
    print('\n' ,'Corte media voos:', media , '\n' , 'corte mediana realizados:',
          mediana , '\n' ,'frameCorteAeroDestino' , '\n' ,frameCorteAeroDestino, '\n')







    print('\n', "********************************************************************",'\n')
    '''
    
    if(ano==0):
        nome='Infraero2015t.xlsx'
    elif(ano==1):
        nome='Infraero2016t.xlsx'
    elif(ano==2):
        nome='Infraero2017t.xlsx'
    elif(ano==3):
        nome='Infraero2018t.xlsx'
    else:
        nome='Infraero2019t.xlsx'
        '''
    nome= 'Infraero' + (str(ano+15)) + '.xlsx'

    with pd.ExcelWriter(nome) as writer:  
        frameGeral.to_excel(writer,sheet_name='frameGeral', index=False) 
        frameEmpresas.to_excel(writer,sheet_name='frameEmpresas', index=False)   
        frameCorteEmpresa.to_excel(writer,sheet_name='frameCorteEmpresa', index=False) 
        frameAeroOrigem.to_excel(writer,sheet_name='frameAeroOrigem', index=False)  
        frameCorteAeroOrigem.to_excel(writer,sheet_name='frameCorteAeroOrigem', index=False)
        frameAeroDestino.to_excel(writer,sheet_name='frameAeroDestino', index=False) 
        frameCorteAeroDestino.to_excel(writer,sheet_name='frameCorteAeroDestino', index=False)

print('\n', "Final dos testes")



