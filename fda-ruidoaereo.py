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

			RUIDO AÉREO
			Hojas de cálculo

"""
A_Senal_Emision_Recepcion = "RuidoAereo/Practica4_puntoA_emision(Ch1) y recepcion(Ch2).xls"
B_Senal_Emision_Recepcion = "RuidoAereo/Práctica4_puntoB_emisión(Ch1) y recepción(Ch2).xls"
C_Senal_Emision_Recepcion = "RuidoAereo/Practica4_puntoC_emisión(Ch1) y recepción(Ch2).xls"

A_Ruido_Emision_Recepcion = "RuidoAereo/Practica4_puntoA_rfondo_emisión(Ch1) y recepción(Ch2).xls"
B_Ruido_Emision_Recepcion = "RuidoAereo/Practica4_puntoB_rfondo_emisión(Ch1) y recepción(Ch2).xls"
C_Ruido_Emision_Recepcion = "RuidoAereo/Practica4_puntoC_rfondo_emisión(Ch1) y recepción(Ch2).xls"

libro_tr_A1 = "Datos tiempo de reverberación/FDA_Práctica3_PtosMedidaRuidoIntenrrumpido/Práctica3_posA_ruido_1/practica1_positionA_ruido1_TR60.xls"
libro_tr_A2 = "Datos tiempo de reverberación/FDA_Práctica3_PtosMedidaRuidoIntenrrumpido/Práctica3_posA_ruido_2/Practica1_posA_ruido_2tr60.xls"
libro_tr_B1 = "Datos tiempo de reverberación/FDA_Práctica3_PtosMedidaRuidoIntenrrumpido/Práctica3_posB_ruido_1/practica1_positionB_ruido1_TR60.xls"
libro_tr_B2 = "Datos tiempo de reverberación/FDA_Práctica3_PtosMedidaRuidoIntenrrumpido/Práctica3_posB_ruido_2/Practica1_posB_ruido_2_tr60.xls"

libro_senal_A = xlrd.open_workbook(A_Senal_Emision_Recepcion)
libro_senal_B = xlrd.open_workbook(B_Senal_Emision_Recepcion)
libro_senal_C = xlrd.open_workbook(C_Senal_Emision_Recepcion)

libro_ruido_A = xlrd.open_workbook(A_Ruido_Emision_Recepcion)
libro_ruido_B = xlrd.open_workbook(B_Ruido_Emision_Recepcion)
libro_ruido_C = xlrd.open_workbook(C_Ruido_Emision_Recepcion)



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
		if dif < 10.0:
			senal_lineal = math.pow(10,0.1*senal[n])
			ruido_lineal = math.pow(10,0.1*ruido[n])
			corr = senal_lineal - ruido_lineal
			corr = 10*(math.log10(corr))
			senal[n] = round(corr,1)

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


def promediado_espacial(A,B,C):

	"""
	Realiza el promediado espacial de 
	los tres puntos, A, B y C.
	"""

	resultado = []

	for n in range(len(A)):
		A_lineal = math.pow(10,0.1*A[n])
		B_lineal = math.pow(10,0.1*B[n])
		C_lineal = math.pow(10,0.1*C[n])
		suma_lineal = A_lineal + B_lineal + C_lineal
		ans =  10*(math.log10((suma_lineal/3)))
		resultado.append(round(ans,1))

	return resultado

def tr_lista(hoja):
	mi_hoja = hoja.sheet_by_index(1)
	return row2list(mi_hoja.row(2)[1:])

def promedio_tr(A1,A2):
	total = []
	for n in range(len(A1)):
		value = float(A1[n])+float(A2[n])/4
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

#def pintar():
"""

PROGRAMA PRINCIPAL

"""

if __name__ == "__main__":


	help = "\r\nUso: python fda.py <emision o recepcion>\r\n"
	help += "Opciones: \r\n"
	help += "\t<emision o recepcion> : puede ser -emision si queremos calcular el nivel de emision\r\n"
	help += "\to -recepcion si queremos calcular el nivel en recepción.\n"
	help += "\tSon obligatorios ambos argumentos.\n"

	

	print("\r\n---Herramienta-fda start--- \r\n\tcc @ Grupo13-FDA\r\n")

	flagDonde = 1
	hoja_senal_A = libro_senal_A.sheet_by_index(flagDonde)
	hoja_senal_B = libro_senal_B.sheet_by_index(flagDonde)
	hoja_senal_C = libro_senal_C.sheet_by_index(flagDonde)

	hoja_ruido_A = libro_ruido_A.sheet_by_index(flagDonde)
	hoja_ruido_B = libro_ruido_B.sheet_by_index(flagDonde)
	hoja_ruido_C = libro_ruido_C.sheet_by_index(flagDonde)

	print("\r\n --- Nivel en emisión en dB ---   \r\n")
	

	A =  [] #Lista niveles en punto A, corregido
	B = [] #Lista niveles en punto B, corregido
	C = [] ##Lista niveles en punto C, corregido

	ruido = []

	"""  Señal A """

	A = calcular(hoja_senal_A)

	"""  Señal B """
		
	B = calcular(hoja_senal_B)

	"""  Señal C """

	C = calcular(hoja_senal_C)

	Lemision = promediado_espacial(A,B,C)

	print Lemision

	flagDonde = 2
	hoja_senal_A = libro_senal_A.sheet_by_index(flagDonde)
	hoja_senal_B = libro_senal_B.sheet_by_index(flagDonde)
	hoja_senal_C = libro_senal_C.sheet_by_index(flagDonde)

	hoja_ruido_A = libro_ruido_A.sheet_by_index(flagDonde)
	hoja_ruido_B = libro_ruido_B.sheet_by_index(flagDonde)
	hoja_ruido_C = libro_ruido_C.sheet_by_index(flagDonde)

	hoja_A1 = xlrd.open_workbook(libro_tr_A1)
	hoja_A2 = xlrd.open_workbook(libro_tr_A2)
	hoja_B1 = xlrd.open_workbook(libro_tr_B1)
	hoja_B2 = xlrd.open_workbook(libro_tr_B2)

	TR15A1 = tr_lista(hoja_A1)
	TR15A2 = tr_lista(hoja_A2)
	TR15B1 = tr_lista(hoja_B1)
	TR15B2 = tr_lista(hoja_B2)

	#print "PRUEBAAAAAAAAAAAAAAAA"
	#print TR15B2

	TR15A = promedio_tr(TR15A1,TR15A2)
	TR15B = promedio_tr(TR15B1,TR15B2)

	print("\r\n --- Nivel en recepción en dB ---   \r\n")

	A = [] #Lista niveles en punto A, corregido
	B = [] #Lista niveles en punto B, corregido
	C = [] ##Lista niveles en punto C, corregido

	"""  Señal A """

	A = calcular(hoja_senal_A)

	"""  Señal B """
		
	B = calcular(hoja_senal_B)

	"""  Señal C """

	C = calcular(hoja_senal_C)

	Lrecepcion = promediado_espacial(A,B,C)

	print Lrecepcion

	print("\r\n --- Ruido en recepción en dB ---   \r\n")


	"""  Ruido A """

	ruidoA = calcular(hoja_ruido_A)

	"""  Ruido B """

	ruidoB = calcular(hoja_ruido_B)

	"""  Ruido C """

	ruidoC = calcular(hoja_ruido_C)

	ruido = promediado_espacial(ruidoA,ruidoB,ruidoC)
	print ruido

	print("\r\n --- Lcorr en dB ---   \r\n")
	Lrecepcion_corr = correccion_ruido(Lrecepcion,ruido)

	print Lrecepcion_corr

	print("\r\n --- Aislamiento D en dB ---   \r\n")

	print calcularD(Lemision,Lrecepcion_corr)

	print("\r\n --- TR60 en segundos ---   \r\n")
	
	

	TR60 = promedio_total_tr(TR15A,TR15B)

	print TR60

	print("\r\n --- DnT en dB ---   \r\n")
	print calcularDnT(TR60[:-2])

	print("\r\n --- R' en dB ---   \r\n")
	print calcularRaparente()

