import math 
import pandas as pd
#==================CARGA TÉRMICA DE RENOVAÇÃO==================================

#calculo da entalpia
def dH(u_amb,u_ext,t_amb,t_ext):
    dh_cv=(1.005*t_ext+u_amb*(2500.9+1.82*t_ext))-(1.005*t_amb+u_amb*(2500.9+1.82*t_amb))
    dh_lt=(1.005*t_amb+u_ext*(2500.9+1.82*t_amb))-(1.005*t_amb+u_amb*(2500.9+1.82*t_amb))
    dh_tot=dh_cv+dh_lt  
    return dh_tot

#calculo da vazão mássica
def v_massica(pes,a,fp,fa):
    ro = 1.225 #[kg/m³] densidade do ar
    v_ef = pes*fp+a*fa #[L/s] vazão efetiva
    v_mass = v_ef*ro/1000 #[kg/s] vazão mássica
    return v_mass

#calculo da carga térmica 
def carga_term_ren(pes,a,fp,fa,u_amb,u_ext,t_amb,t_ext):
    dh = dH(u_amb,u_ext,t_amb,t_ext)
    m_ren = v_massica(pes,a,fp,fa) 
    c_term=m_ren*dh*1000
    return c_term


    



#==================RADIAÇÕES===================================================
#Calculo do angulo solar
def angSol(hSol): 
    w = 15*(hSol -12)
    return w

#Calculo de declinação
def declin(nDia):
    dec=[]
    for i in range(nDia.size):
        value = 23.5*math.sin(math.radians(360*(284+nDia[i])/365))
        dec.append(value)
    return dec

#Calculo do cosseno do angulo de incidencia
def cosTheta(lat,incli,aziSup,dec,angSol):
    a = math.sin(lat)*math.cos(incli)
    b = math.cos(lat)*math.sin(incli)*math.cos(aziSup)
    c = math.sin(incli)*math.sin(aziSup)
    d = math.cos(lat)*math.cos(incli)
    e = math.sin(lat)*math.sin(incli)*math.cos(aziSup)
    cos_theta = []
    for i in range(dec.size):
        value = (a-b)*math.sin(dec[i])+(c*math.sin(angSol[i])+(d+e)*math.cos(angSol[i]))*math.cos(dec[i])
        cos_theta.append(value)
    return cos_theta    

#Calculo da altura solar
def altSol(dec,lat,angSol):
    a_Sol =[]
    for i in range(dec.size):
        value = math.asin( math.sin(math.radians(dec[i]))*math.sin(math.radians(lat))+math.cos(math.radians(dec[i]))*math.cos(math.radians(lat))*math.cos(math.radians(angSol[i])))
        a_Sol.append(value)
    return a_Sol

#Calculo radiação direta
def radDir(aSol,file,mes):
    cons = pd.read_excel(file)
    Id=[]
    for i in range(aSol.size):
        n =  mes[i]
        a = cons['A'][n-1]
        b = cons['B'][n-1]
        value = a/(math.e**(b/(math.sin(aSol[i]))))
        Id.append(value)
    return Id    

#Calculo radiação difusa
def radDif(Id,dec,ro,file,aSol,mes):
    cons = pd.read_excel(file)
    If = []
    for i in range(Id.size):
        n =  mes[i]
        c = cons['C'][n-1]
        Fg = 0.5*(1-math.cos(math.radians(dec[i]))) 
        Fs = 0.5*(1+math.cos(math.radians(dec[i]))) 
        Is = c*Id[i]*Fs
        Ir = Id[i]*(c+math.sin(math.radians(aSol[i])))*Fg*ro
        value = Is+Ir
        If.append(value)
    return If


#=======================CARGA TÉRMICA ENVOLTÓRIA===============================
#calculo da radiação total na superficie
def radTotSup(cosTheta,Id,If):
    I=[]
    for i in range(Id.size):
        a = math.degrees(math.acos(cosTheta[i]))
        
        if abs(a)<90:
            value = (Id[i] * cosTheta[i]) + If[i]
        else:
            value = If[i]
        I.append(value)     
    return I

#Calculo da temperatura Sol-Ar
def tempSolAr(te,rse,alpha,I,opaca):
    
    if opaca =='Sim' or opaca =='sim':
        t_Sol_Ar = te+rse*alpha*I
    else:
        t_Sol_Ar = 0
    return t_Sol_Ar

#Calculo da carga térmica convectiva
def cTermConv(area,U,f,t_amb,t_ext,t_med,lag,opaca):
    
    t_theta = []
    for i in range(t_ext.size):
        if i<5:
            n = ((t_ext.size)-lag-1-i)
            value = t_ext[n]
            
        else:
            value = t_ext[(i-lag)]
        t_theta.append(value)
    if opaca == 'Sim' or opaca == 'sim':
        c_term_conv = area*U*(t_med-t_amb)+area*U*(t_theta-t_med)*f 
    else:
        c_term_conv = area*U*(t_ext-t_amb)
    return c_term_conv
       
#Calculo da carga térmica radiante
def cTermRad(area,Fs,I,U,t_ext,t_amb,opaca):
    if opaca == 'Sim' or opaca == 'sim':
        c_term_rad = 0
    else:
        c_term_rad = area*(Fs*I+U*(t_ext-t_amb))
    return c_term_rad



#==============================================================================
