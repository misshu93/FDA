#!/usr/bin/python
# -*- coding: utf-8 -*-

import os
import sys,xlrd
import math
import commands

"""
Script para la práctica 3 de FDA
Seleccionar al inicio si se quiere calcular emision o recepción

"""


"""

			RUIDO IMPACTOS
			Hojas de cálculo

"""
A_Senal_Recepcion = "RuidoImpacto/Practica1_mip1_micros1-23.xls"
B_Senal_Recepcion = "RuidoImpacto/Practica1_mip1_micros3-40.xls"
C_Senal_Recepcion = "RuidoImpacto/Practica1_mip4_micros3-40.xls"
D_Senal_Recepcion = "RuidoImpacto/Practica1_mip4_micros1-20.xls"
E_Senal_Recepcion = "RuidoImpacto/Practica1_mip2_micros1-21.xls"
F_Senal_Recepcion = "RuidoImpacto/Practica1_mip3_micros1-20.xls"
G_Senal_Recepcion = "RuidoImpacto/Practica1_mip2_micros3-40.xls"
H_Senal_Recepcion = "RuidoImpacto/Practica1_mip3_micros3-40.xls"


libro_senal_A = xlrd.open_workbook(A_Senal_Recepcion)
libro_senal_B = xlrd.open_workbook(B_Senal_Recepcion)
libro_senal_C = xlrd.open_workbook(C_Senal_Recepcion)
libro_senal_D = xlrd.open_workbook(D_Senal_Recepcion)
libro_senal_E = xlrd.open_workbook(E_Senal_Recepcion)
libro_senal_F = xlrd.open_workbook(F_Senal_Recepcion)
libro_senal_G = xlrd.open_workbook(G_Senal_Recepcion)
libro_senal_H = xlrd.open_workbook(H_Senal_Recepcion)

A_Ruido_Recepcion = "RuidoImpacto/Practica1_mip1_micros1-2_rfondo0.xls"
B_Ruido_Recepcion = "RuidoImpacto/Practica1_mip1_micros3-4_rfondo0.xls"
C_Ruido_Recepcion = "RuidoImpacto/Practica1_mip4_micros3-4_rfondo0.xls"
D_Ruido_Recepcion = "RuidoImpacto/Practica1_mip4_micros1-2_rfondo0.xls"
E_Ruido_Recepcion = "RuidoImpacto/Practica1_mip2_micros1-2_rfondo1.xls"
F_Ruido_Recepcion = "RuidoImpacto/Practica1_mip3_micros1-2_rfondo0.xls"
G_Ruido_Recepcion = "RuidoImpacto/Practica1_mip2_micros3-4_rfondo0.xls"
H_Ruido_Recepcion = "RuidoImpacto/Practica1_mip3_micros3-4_rfondo0.xls"


libro_Ruido_A = xlrd.open_workbook(A_Ruido_Recepcion)
libro_Ruido_B = xlrd.open_workbook(B_Ruido_Recepcion)
libro_Ruido_C = xlrd.open_workbook(C_Ruido_Recepcion)
libro_Ruido_D = xlrd.open_workbook(D_Ruido_Recepcion)
libro_Ruido_E = xlrd.open_workbook(E_Ruido_Recepcion)
libro_Ruido_F = xlrd.open_workbook(F_Ruido_Recepcion)
libro_Ruido_G = xlrd.open_workbook(G_Ruido_Recepcion)
libro_Ruido_H = xlrd.open_workbook(H_Ruido_Recepcion)


def calcular_promedio(lista_valores):
    """
    Calcula el promedio logaritimico de 
    una lista de valores en formato x,x.
    """
    valor_acumulado = 0.0
    for valor in lista_valores:
        lista_valor = valor.split(",")
        valor = ".".join(lista_valor)
        valor = float(valor)
        valor_elevado = math.pow(10,0.1*valor)
        valor_acumulado += float(valor_elevado)
    long_lista = float(len(lista_valores))
    resultado = 10*(math.log10(valor_acumulado/long_lista))
    return round(resultado,1)


def cell2str(cell):

	"""
	Pasa el valor de cada celda 
	del libro de Excel a string.
	"""

 	return str(cell).split(":")[1]

def row2list(row):

	"""
	Pasa una columna de Excel a lista.
	"""

	lista = []
	for cell in row:
		value = cell2str(cell)
		lista.append(value)

	return lista
		

def correccion_ruido(senal,ruido):

	"""
	Corrige el ruido si la diferencia entre
	la señal y el ruido es < de 10dB

	"""

	for n in range(len(senal)):
		dif = senal[n]-ruido[n]
		if dif < 10 and dif > 6:
			senal_lineal = math.pow(10,0.1*senal[n])
			ruido_lineal = math.pow(10,0.1*ruido[n])
			corr = senal_lineal - ruido_lineal
			corr = 10*(math.log10(corr))
			senal[n] = round(corr,1)
		elif dif <= 6:
			senal[n] = round(senal[n] - 1.3,1)

	return senal

def calcular(hoja_senal):
	"""
	Saca las columnas de cada libro,
	lo pasa a lista.

	"""
	lista = []

	for columna in range(3,19):

		columna = hoja_senal.col(columna)[2:]
		columna = row2list(columna)
		lista.append(calcular_promedio(columna))

	return lista


def promediado_espacial(A,B,C,D,E,F,G,H):

	"""
	Realiza el promediado espacial de 
	todos los puntos
	"""

	resultado = []

	for n in range(len(A)):
		A_lineal = math.pow(10,0.1*A[n])
		B_lineal = math.pow(10,0.1*B[n])
		C_lineal = math.pow(10,0.1*C[n])
		D_lineal = math.pow(10,0.1*D[n])
		E_lineal = math.pow(10,0.1*E[n])
		F_lineal = math.pow(10,0.1*F[n])
		G_lineal = math.pow(10,0.1*G[n])
		H_lineal = math.pow(10,0.1*H[n])

		suma_lineal = A_lineal + B_lineal + C_lineal + D_lineal + E_lineal + F_lineal + G_lineal + H_lineal
		ans =  10*(math.log10((suma_lineal/8)))
		resultado.append(round(ans,1))

	return resultado

def tr_lista(hoja):
	mi_hoja = hoja.sheet_by_index(1)
	return row2list(mi_hoja.row(2)[1:])

def promedio_tr(A1,A2):
	total = []
	for n in range(len(A1)):
		value = float(A1[n])+float(A2[n])/2
		total.append(round(value,1))

	return total

def promedio_total_tr(A,B):
	total = []

	for n in range(len(A)):
		valor = (A[n]+B[n])/2
		total.append(round(valor,1))

	return total
def calcularDnT(TR60):
	D = [27.5, 26.4,27,27.2,29.4,29.6,27.2,29.2,30.2,31.9,31.3,30.5,30.8,30.5,26.6,25.5]
	DnT = []
	for n in range(len(D)):
		valor = D[n] + 10*math.log10(TR60[n]/0.5)
		DnT.append(round(valor,1))
	return DnT

def calcularRaparente():
	D = [27.5, 26.4,27,27.2,29.4,29.6,27.2,29.2,30.2,31.9,31.3,30.5,30.8,30.5,26.6,25.5]
	Raparente = []
	for n in range(len(D)):
		valor = D[n] + 10*math.log10(12.3/10)
		Raparente.append(round(valor,1))


	return Raparente

def calcularD(Lemision,Lrecepcion):
	D = []
	for n in range(len(Lemision)):
		valor =Lemision[n]-Lrecepcion[n]
		D.append(round(valor,1))
	return D

def calcular_A():
	A = []
	V = 10.4*9.44*3
	TR60 = [0.9, 0.9, 0.5, 0.4, 0.6, 0.6, 0.6, 0.8, 0.8, 0.7, 0.7, 0.7, 0.7, 0.8, 0.8, 0.8, 0.8, 0.8]
	for n in range(len(TR60)):
		value = 0.161*V*TR60[n]
		A.append(round(value,1))

	return A

def calcular_Ln(Li,A):

	Ln = []
	for n in range(len(Li)):
		value = Li[n] + 10*math.log10(A[n]/10)
		Ln.append(round(value,1))

	return Ln


def calcular_LnT(Li):

	TR60 = [0.9, 0.9, 0.5, 0.4, 0.6, 0.6, 0.6, 0.8, 0.8, 0.7, 0.7, 0.7, 0.7, 0.8, 0.8, 0.8]

	LnT = []
	for n in range(len(Li)):
		value = Li[n] - 10*math.log10(TR60[n]/0.5)
		LnT.append(round(value,1))

	return LnT
"""

PROGRAMA PRINCIPAL

"""

if __name__ == "__main__":

	print("\r\n---Herramienta-fda start--- \r\n\tcc @ Grupo13-FDA\r\n")
	"""
	SEÑAL
	"""

	flagDonde = 1
	hoja_senal_A = libro_senal_A.sheet_by_index(flagDonde)
	hoja_senal_B = libro_senal_B.sheet_by_index(flagDonde)
	hoja_senal_C = libro_senal_C.sheet_by_index(flagDonde)
	hoja_senal_D = libro_senal_D.sheet_by_index(flagDonde)
	hoja_senal_E = libro_senal_E.sheet_by_index(flagDonde)
	hoja_senal_F = libro_senal_F.sheet_by_index(flagDonde)
	hoja_senal_G = libro_senal_G.sheet_by_index(flagDonde)
	hoja_senal_H = libro_senal_H.sheet_by_index(flagDonde)
	

	A = calcular(hoja_senal_A)
	B = calcular(hoja_senal_B)
	C = calcular(hoja_senal_C)
	D = calcular(hoja_senal_D)
	E = calcular(hoja_senal_E)
	F = calcular(hoja_senal_F)
	G = calcular(hoja_senal_G)
	H = calcular(hoja_senal_H)


	L1 = promediado_espacial(A,B,C,D,E,F,G,H)

	flagDonde = 2
	hoja_senal_A = libro_senal_A.sheet_by_index(flagDonde)
	hoja_senal_B = libro_senal_B.sheet_by_index(flagDonde)
	hoja_senal_C = libro_senal_C.sheet_by_index(flagDonde)
	hoja_senal_D = libro_senal_D.sheet_by_index(flagDonde)
	hoja_senal_E = libro_senal_E.sheet_by_index(flagDonde)
	hoja_senal_F = libro_senal_F.sheet_by_index(flagDonde)
	hoja_senal_G = libro_senal_G.sheet_by_index(flagDonde)
	hoja_senal_H = libro_senal_H.sheet_by_index(flagDonde)



	A = calcular(hoja_senal_A)
	B = calcular(hoja_senal_B)
	C = calcular(hoja_senal_C)
	D = calcular(hoja_senal_D)
	E = calcular(hoja_senal_E)
	F = calcular(hoja_senal_F)
	G = calcular(hoja_senal_G)
	H = calcular(hoja_senal_H)


	L2 = promediado_espacial(A,B,C,D,E,F,G,H)



	print("\r\n --- Nivel promediado en recepción en dB ---   \r\n")

	senal = promedio_total_tr(L1,L2)

	print senal


	"""
	RUIDO
	"""

	flagDonde = 1
	hoja_Ruido_A = libro_Ruido_A.sheet_by_index(flagDonde)
	hoja_Ruido_B = libro_Ruido_B.sheet_by_index(flagDonde)
	hoja_Ruido_C = libro_Ruido_C.sheet_by_index(flagDonde)
	hoja_Ruido_D = libro_Ruido_D.sheet_by_index(flagDonde)
	hoja_Ruido_E = libro_Ruido_E.sheet_by_index(flagDonde)
	hoja_Ruido_F = libro_Ruido_F.sheet_by_index(flagDonde)
	hoja_Ruido_G = libro_Ruido_G.sheet_by_index(flagDonde)
	hoja_Ruido_H = libro_Ruido_H.sheet_by_index(flagDonde)



	A = calcular(hoja_Ruido_A)
	B = calcular(hoja_Ruido_B)
	C = calcular(hoja_Ruido_C)
	D = calcular(hoja_Ruido_D)
	E = calcular(hoja_Ruido_E)
	F = calcular(hoja_Ruido_F)
	G = calcular(hoja_Ruido_G)
	H = calcular(hoja_Ruido_H)


	ruido1 = promediado_espacial(A,B,C,D,E,F,G,H)

	flagDonde = 2
	hoja_Ruido_A = libro_Ruido_A.sheet_by_index(flagDonde)
	hoja_Ruido_B = libro_Ruido_B.sheet_by_index(flagDonde)
	hoja_Ruido_C = libro_Ruido_C.sheet_by_index(flagDonde)
	hoja_Ruido_D = libro_Ruido_D.sheet_by_index(flagDonde)
	hoja_Ruido_E = libro_Ruido_E.sheet_by_index(flagDonde)
	hoja_Ruido_F = libro_Ruido_F.sheet_by_index(flagDonde)
	hoja_Ruido_G = libro_Ruido_G.sheet_by_index(flagDonde)
	hoja_Ruido_H = libro_Ruido_H.sheet_by_index(flagDonde)

	
	A = calcular(hoja_Ruido_A)
	B = calcular(hoja_Ruido_B)
	C = calcular(hoja_Ruido_C)
	D = calcular(hoja_Ruido_D)
	E = calcular(hoja_Ruido_E)
	F = calcular(hoja_Ruido_F)
	G = calcular(hoja_Ruido_G)
	H = calcular(hoja_Ruido_H)


	ruido2 = promediado_espacial(A,B,C,D,E,F,G,H)


	print("\r\n --- Ruido promediado en recepción en dB ---   \r\n")
	

	ruido = promedio_total_tr(ruido1,ruido2)
	print ruido


	print("\r\n --- Área de absorción sonora, en función de la frecuencia (m2) ---   \r\n")

	area = calcular_A()

	print area

	print("\r\n --- Li (dB) ---   \r\n")

	Li = correccion_ruido(senal,ruido)

	print Li

	print("\r\n --- L'n (dB) ---   \r\n")

	print calcular_Ln(Li,area)

	print("\r\n --- L'nT (dB) ---   \r\n")

	print calcular_LnT(Li)

