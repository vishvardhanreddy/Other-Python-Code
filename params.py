'''
Created on Aug 6, 2014

@author: vvaka
'''
import os
import sys
import re
import xlrd
import math
from time import strftime
import matplotlib
matplotlib.use('Agg')
import numpy as np
import matplotlib.pyplot as plt
import pylab as pl
import csv

def is_number(s):
# This function checks if the number is passed is number or character, if number returns True, else it will
# return False

	try:
        	float(s)
        	return True
    	except ValueError:
        	return False

def getDcuList():
    
	dcuList = {
     		'DCU-100':'2.1',
     		'DCU-350':'3.0',
     		'DCU-450':'3.5',
     		'DCU-550':'3.9',
     		'DCU-750':'5.0',
     		'DCU-950':'5.5',
     		'DCU-1150':'6.2',
     		'DCU-1350':'6.4',
     		'DCU-1550':'7.2',
     		'DCU-1950':'8.8',
     		'DCU-E-200':'5.5',
     		'DCU-E-350':'7.0',
     		'DCU-L-300':'3.0',
     		'DCU-L-600':'4.2',
     		'DCU-L-700':'4.6',
     		'DCU-L-800':'5.0',
     		'DCU-L-1000':'5.8',
     		'DCU-L-1100':'6.0',
        	}
    	return dcuList

def getNodeCount(processdir):
        
	inputFile = processdir+'/WDMANS_OTS-OCH_TH.csv'
	try:
		inFile = open(inputFile,'r')
        
	# Opening input file   
	except IOError:
        	print "[ERROR] Cound't open file: " + inputFile+". Failed to generate Node Count."
        	return None
    	nodeList = []
	for line in inFile:
		line = re.sub('[\n\r]','',line)
		flds = line.split(',')
		if flds[0] not in nodeList:
			nodeList.append(flds[0])
	nodeList = set(nodeList)
	if '#Node' in nodeList:
		nodeList.remove('#Node')
	nodeList = filter(None, nodeList)
        nodeCount = len(nodeList)-1
	return nodeCount

def setDHIFlag(DHI):
    
	if DHI >= 4.00:
        	DHIFlag = 'GREEN'
    	elif (DHI < 4.00 and DHI >= 2.00):
        	DHIFlag = 'YELLOW'
    	elif (DHI < 2.00):
        	DHIFlag = 'RED'
    	return DHIFlag

def createTable21(reportsdir, processdir):
    
	inputFile = processdir+'/WDMANS_OTS-OCH_TH.csv'
	outputFile = reportsdir+'/Table2.1.csv'
    	greenDHI = 0
    	yellowDHI = 0
    	redDHI = 0
    	oldNodeName = 'X'
    	DHITotal = 0
        
    	try:
        	inFile = open(inputFile,'r')
        # Opening input file   
    	except IOError:
        	print "[ERROR] Cound't open file: " + inputFile +". Failed to generate Table2.1.csv"
        	return None
    
    	try:
        	outFile = open(outputFile,'w')
        	# Open output file
    	except IOError:
        	print "[ERROR] Cound't write"+outputFile
        	return False
    
  
  
    	outFile.write('Node Name, Shelf/Slot/Port, Equipment, Origin, ANS Parameter, ANS Value, Card Reading, DHI\n')
    	
	for line in inFile:
        	line = re.sub('[\n\r]','',line)
		flds = line.split(',')
		
		if (is_number(flds[9])) and (is_number(flds[10])):
            # if there are values for both WDM-ANS and Actual Card, DHI is the difference between them  
            		DHI = float(flds[9]) - float(flds[10])  
            		DHITotal = DHI+DHITotal
            # if only WDM-ANS value is available, then DHI is WDM-ANS

      
            # Since we are only picking Yellow and Red DHI, we are picking absolute value of DHI only if it is
            # not equal to 0
            		DHI = math.fabs(DHI)
            		if DHI == 0.0:
                		greenDHI  = greenDHI + 1   
            		if (DHI != 0.0): 
                		if DHI > 1:
                    			DHIColor = 'RED'
                    			redDHI = redDHI + 1
                    			format.set_bg_color('#FF0000')
                		elif DHI <= 1:
                    			DHIColor = 'YELLOW'
                    			yellowDHI = yellowDHI + 1    
                    			format.set_bg_color('#FFFF00')
                		StrinDHI = str(DHI)
                		row2 = str(int(flds[1]))+'/'+str(int(flds[2]))+'/'+str(int(flds[8]))
                		outFile.write(flds[0]+','+row2+','+flds[4]+','+flds[6]+','+flds[7]+','+flds[9]+','+flds[10]+','+DHIColor+'\n')
  
  	if (yellowDHI+redDHI) == 0:
        	outFile.write('N/A,N/A,N/A,N/A,N/A,N/A,N/A,N/A,N/A\n')
 
	totalNo = redDHI+yellowDHI+greenDHI
    	total = (redDHI*1)+(yellowDHI*3)+(5*greenDHI) 
    	averageDHI = total/float(totalNo)
    	#print averageDHI
    	outFile.close()     
    	return averageDHI

def createTable22(reportsdir, processdir):

 
	inputFile = processdir+'/WDMANS_OTS-OCH_TH.csv'
        outputFile = reportsdir+'/Table2.2.csv'
    	try:
        	inFile = open(inputFile,'r')
		#workbook = xlrd.open_workbook(inputFile)
        # Opening input file
    	except IOError:
        	print "[ERROR] Cound't open file: " + inputFile+". Failed to generate Table2.2.csv"
        	return None
    
    	try:
        	outFile = open(outputFile,'w')
        	# Open output file
    	except IOError:
        	print "[ERROR] Cound't write"+outputFile
        	return False
  
    	greenDHI = 0
    	yellowDHI = 0
    	redDHI = 0  
    	# Reading all the worksheet names  
  
	outFile.write('Node Name, Shelf/Slot/Port, Equipment, Origin, ANS Parameter, NE Value, CTP Value, DHI\n')
  	
	for line in inFile:
		line = re.sub('[\n\r]','',line)
		flds = line.split(',')

        	if (is_number(flds[13])) and (flds[13] != 0):
            	# Column N values in file are DHI values between Card level readings and CTP are captured in row[13] 
            		DHI = math.fabs(float(flds[13]))
            		if DHI == 0.0:
                		greenDHI = greenDHI + 1
            		if (DHI != 0.0):
                		if DHI > 1:
                    			DHIColor = 'RED'
                    			redDHI = redDHI + 1
                		elif DHI <= 1:
                    			DHIColor = 'YELLOW' 
                    			yellowDHI = yellowDHI + 1
                		port = int(flds[8])
                		port = str(port)
                		#print port 
                		row2 = str(int(flds[1]))+'/'+str(int(flds[2]))
                		outFile.write(flds[0]+','+row2+'/'+port+','+flds[4]+','+flds[6]+','+flds[7]+','+flds[9]+','+flds[12]+','+DHIColor+'\n')

	if (yellowDHI+redDHI) == 0:
        	outFile.write('N/A,N/A,N/A,N/A,N/A,N/A,N/A,N/A,N/A\n')

    	totalNo = redDHI+yellowDHI+greenDHI
    	total = (redDHI*1)+(yellowDHI*3)+(5*greenDHI) 

    	averageDHI = total/float(totalNo)
    	return averageDHI
    	outFile.close()      

def createTable23(reportsdir, processdir):


    	#outputFile = 'Table2.3.csv'
    	# Second argument is considered to be Output file name
        inputFile = processdir+'/WDMANS_OTS-OCH_TH.csv'
        outputFile = reportsdir+'/Table2.3.csv'

    	yellow = 0
    	red = 0
    	oldNodeName = 'X'
    	try:
       		# Opening input file
		inFile = open(inputFile,'r')
    	except IOError:
        	print "[ERROR] Cound't open file: " + inputFile +". Failed to generate Table2.3.csv"
        	return None
    
    	try:
        	outFile = open(outputFile,'w')
        # Open output file
    	except IOError:
        	print "[ERROR] Cound't write"+outputFile
        	return False
    
  	# Reading all the worksheet names  
    	greenDHI = 0
    	yellowDHI = 0
    	redDHI = 0
    
   	outFile.write('Node Name, Shelf/Slot/Port, Equipment, Origin, Parameter, Power (dBm), CTP LOS Threshold, DHI\n')
	for line in inFile:
	# Column P values in file are DHI values between Card level readings and CTP Los Threshold are captured in row[13] 
		line = re.sub('[\n\r]','',line)
		flds = line.split(',')
		if ('Site' in flds[0]) and (len(flds)>=15):
			if (is_number(flds[15])) and flds[12] and flds[14]:
            			if 'LOS Threshold' in flds[7]:
            			# Since we are only looking for Yellow and Red DHI values, we only capture values below 1db
                			if (float(flds[15]) >= 1.0):
                    				DHI = 'GREEN'
                    				greenDHI = greenDHI + 1
                
                			elif (float(flds[15])) < 0.0:
                    				DHI = 'RED'
                    				redDHI = redDHI + 1
                			else:
                    				DHI = 'YELLOW'
                    				yellowDHI = yellowDHI + 1
                			if DHI != 'GREEN':
                    				row2 = str(int(flds[1]))+'/'+str(int(flds[2]))+'/'+str(int(flds[8]))
                    				outFile.write(flds[0]+','+row2+','+flds[4]+','+flds[6]+','+flds[7]+','+flds[14]+','+flds[12]+','+DHI+'\n')
  

    	totalNo = redDHI+yellowDHI+greenDHI
    	total = (redDHI*1)+(yellowDHI*3)+(5*greenDHI) 
    	average = total/float(totalNo)
    	return average   
    	outFile.close()     

def createTable24(reportsdir, processdir):
    
    	inputFile = processdir+'/WDMANS_OTS-OCH_TH.csv'
        outputFile = reportsdir+'/Table2.4.csv'
        sidefile = processdir+'/zRTRV-WDMSIDE.csv'
	yellow = 0
    	red = 0
    	oldNodeName = 'X'
    	greenDHI = 0
    	yellowDHI = 0
    	redDHI = 0
    	try:
  		inFile = open(inputFile,'r') 
    	except IOError:
        	print "[ERROR] Cound't open file: " + inputFile +". Failed to generate Table2.4.csv"
        	return None
       
    

    	try:
        	outFile = open(outputFile,'w')
  
    	except IOError:
        	print "[ERROR] Cound't write"+outputFile
        	return False
    
    
    	outFile.write('Node Name,Side,Expected Min,Expected Max,Actual Loss,DHI\n')
	for line in inFile:
		line = re.sub('[\n\r]','',line)
		flds = line.split(',')
		 	 
        	if ('Min Expected Span Loss' in (flds[7])) and (float(flds[9] != 'N/A')):
            		aValue = float(flds[9])
            		ActualLoss = str(flds[14])
            		ActualLoss = ActualLoss.strip('(ACTUAL SPAN LOSS)/')
            		ActualValue = ActualLoss
      
        	elif ('Max Expected Span Loss' in (flds[7])) and (float(flds[9] != 'N/A')):
            		bValue = float(flds[9])     
      
            		C = bValue - aValue
            		ThresholdCenterA = (aValue+(C*0.25))
            		ThresholdCenterB = (bValue-(C*0.25))
            		if ActualLoss:
                		if  ActualValue < ThresholdCenterA:
                    			DHI = 'GREEN'
                    			greenDHI = greenDHI + 1
                		elif ActualValue >= ThresholdCenterB:
                    			DHI = 'YELLOW'
                    			yellowDHI = yellowDHI + 1
                		else:
                    			DHI = 'RED'
                    			redDHI = redDHI + 1
                		LineIn = 'LINE-'+str(int(flds[1]))+'-'+str(int(flds[2]))+'-'+str(flds[3])
                		side = 'N/A'
                
                		try:   
                    			sideFile = open(sidefile,'r')
       
                		except IOError:
                    			print "[ERROR] Cound't open _pRTRV-WDMSIDe.csv"
                    			return False  

                		for line in sideFile:    
                    			line = re.sub('[\n\r]','',line)    
                    			sflds = line.split(',')

                    			if (LineIn == str(sflds[3])) and (str(flds[0]) == str(sflds[0])):
                        			side = sflds[2]
                		sideFile.close()        
                		row2 = str(int(flds[1]))+'/'+str(int(flds[2]))+'/'+str(int(flds[8]))    
                		outFile.write(flds[0]+','+side+','+str(aValue)+','+str(bValue)+','+ActualLoss+','+DHI+'\n')
            		aValue = 'N/A'
            		ActualLoss = 'N/A'

    	totalNo = redDHI+yellowDHI+greenDHI
    	total = (redDHI*1)+(yellowDHI*3)+(5*greenDHI) 
    	average = total/float(totalNo)
    	return average   
    	outFile.close()     


def createFig21(reportsdir, processdir):
  
    	inputFile = processdir+'/WDMANS_OTS-OCH_TH.csv'
	outputFile = reportsdir+'Figure2.1.csv'
    	try:
        	outFile = open(outputFile,'w')
  
    	except IOError:
        	print "[ERROR] Cound't write"+outputFile
        	return False  

    	yellow = 0
    	red = 0
    	green = 0
    	greenCTP = 0
    	yellowCTP = 0
    	redCTP = 0
    	greenCTPLos = 0
    	yellowCTPLos = 0
    	redCTPLos = 0
    	oldNodeName = 'X'
    	try:
  		inFile = open(inputFile,'r') 
    	except IOError:
        	print "[ERROR] Cound't open file: " + inputFile +". Failed to generate Figure 2.1"
        	return None
    
    	try:
        	outFile = open(outputFile,'w')
    	except IOError:
        	print "[ERROR] Cound't write"+outputFile
        	return False

  # defining bold format
      
    	outFile.write('Green,Yellow,Red\n')
    	green = 0
    	yellow = 0
    	red = 0
        for line in inFile:
		line = re.sub('[\n\r]','',line)
		row = line.split(',')
		
		if ('Site' in row[0]):
        		nodeName = row[0]
   
        		if (is_number(row[9])) and (is_number(row[10])):
            			DHI21 = float(row[9]) - float(row[10])  
            			DHI21 = math.fabs(DHI21)
         #outFile.write(nodeName+',Green,Yellow,Red'+'\n')
            			if (DHI21 == 0.0):    
                			green = green +1

            			elif (DHI21 <= 1.0):
                			yellow = yellow +1

            			elif (DHI21 > 1.0):
                			red = red+1
    
        		elif (is_number(row[13])):
          
            			if ((float(row[13])) == 0):    
                			green = green +1
            			elif ((float(row[13])) <= 1.0):
                			yellow = yellow +1
            			elif ((row[13]) > 1.0):
                			red = red+1
        
        		elif (len(row)>=15) and (is_number(row[15])):
            			if ((float(row[15])) >= 1.0):
                			green = green + 1
            			elif ((float(row[15])) >= 0.0):
                			yellow = yellow + 1
            			elif ((float(row[15])) < 0.0):
                			red = red + 1
          
  
    	totalANSSetPoints = green + yellow + red 
    	average = (green+yellow+red)/3
    	outFile.write(str(green)+','+str(yellow)+','+str(red)+'\n')    
    	outFile.close()
        os.chdir(reportsdir) 
    	fig = pl.figure()
    	ax = pl.subplot()
    	y = [green,yellow,red]
   
    	width = 1.0
    	x = ["Green","Yellow","Red"]	
    	ax.bar(range(len(y)),y)
    	ax.text(0.4,green-10,green,rotation=0, va = 'bottom')
    	ax.text(1.4,yellow-10,yellow,rotation=0, va = 'bottom')
    	ax.text(2.4,red-10,red,rotation=0, va = 'bottom')
    
    	ind = np.arange(len(y))
    
    	barlist = plt.bar(ind, y)
    	barlist[0].set_color('g')
    	barlist[1].set_color('y')
    	barlist[2].set_color('r')
    	plt.xticks(ind + width/2.4, x)
    
    	ax.set_xticklabels(x, rotation = 45)
    	fig.autofmt_xdate()
    	pl.show(fig)	
    	pl.savefig("figure2.1.png")
	os.chdir('..')

def createFig31(Green, Yellow, Red, reportsdir):

	os.chdir(reportsdir)
    	fig = pl.figure()
    	ax = pl.subplot()
    	y = [Green, Yellow, Red]
    	x = ["Green","Yellow","Red"]
    	ax.bar(range(len(y)),y)
    	ax.text(0.4,Green-2,Green,rotation=0, va = 'bottom')
    	ax.text(1.4,Yellow-2,Yellow,rotation=0, va = 'bottom')
    	ax.text(2.4,Red-2,Red,rotation=0, va = 'bottom')
    	width = 1.0

    	ind = np.arange(len(y))
    	barlist = plt.bar(ind, y)
    	barlist[0].set_color('g')
    	barlist[1].set_color('y')
    	barlist[2].set_color('r')
    	plt.xticks(ind + width/2.4, x)

    	ax.set_xticklabels(x, rotation = 45)
    	fig.autofmt_xdate()
    	pl.show(fig)
    	pl.savefig("figure3.1.png")
	os.chdir('..')

 
def createTable31(reportsdir, processdir):

        outputFile = reportsdir+'/Table3.1.csv'

    	try:
        	outFile = open(outputFile,'w')
  
    	except IOError:
        	print "[ERROR] Cound't write"+outputFile
        	return False  

    	dcuDict = getDcuList()
    	outFile.write('DCU Type, Loss (dB)\n')
    	for key, value in dcuDict.items():
        	outFile.write(str(key)+','+str(float(value))+'\n')
    	outFile.close()     
    	return None
    
def createTable32(reportsdir, processdir):

 	inputFile = processdir+'/Optical_Power_Levels.csv'
        outputFile = reportsdir+'/Table3.2.csv'
	#dcuFile = 'Process/Table3.1.csv'
    	try:
        	inFile = open(inputFile,'r')
    	except IOError:
        	print "[ERROR] Cound't open file: " + inputFile +". Failed to generate Table 3.2.csv"
        	return None
    
    	try:
        	outFile = open(outputFile,'w')
    	except IOError:
        	print "[ERROR] Cound't write"+outputFile
        	return None
    
    	dictDCU = getDcuList()
    
    	outFile.write('Tx Node, Shelf/Slot/Port, Equipment, Rx Node, Shelf/Slot/Port, Equipment, Power Delta (dB), Attenuator Loss, DCU Loss\n')  
    	redDHI = 0
    	yellowDHI = 0
    	greenDHI = 0
    	for line in inFile:    
        	line = re.sub('[\n\r]','',line)    
        	flds = line.split(',')
        	ATTValue = ''
        	dcuValue = ''
        	if (is_number(flds[6])) and ((math.fabs(float(flds[6]))) > 2.0):
            		redDHI = redDHI + 1
            		if flds[13]:
                		if 'ATT' in str(flds[12]):
                    			ATTValue = str(flds[12])
                    			ATTValue = ATTValue[-2:]
                		elif 'DCU' in str(flds[12]):
                   			dcuValue = dictDCU[str(flds[12])]
                
                		if flds[8]:
                    			rxNode = flds[8]
                		else:
                    			rxNode = flds[0]
                    
                		outFile.write(flds[0]+','+flds[4]+','+flds[3]+','+rxNode+','+flds[11]+','+flds[10]+','+flds[6]+','+ATTValue+','+dcuValue+'\n')    
        
        	elif (is_number(flds[6])) and ((math.fabs(float(flds[6]))) > 1.0):
            		yellowDHI = yellowDHI + 1
        	elif (is_number(flds[6])) and ((math.fabs(float(flds[6]))) <= 1.0):
            		greenDHI = greenDHI + 1
            
    	outFile.close()     
    	inFile.close()
    	DHI = redDHI + yellowDHI + greenDHI
    	DHI = ((redDHI) + (yellowDHI*3) +(greenDHI*5))/DHI
    	createFig31(greenDHI, yellowDHI, redDHI, reportsdir)
    	return DHI

def createTable33(reportsdir, processdir):
  
    	#outputFile = 'Table3.3.csv'
  	inputFile = processdir+'/Optical_Power_Levels.csv'
        outputFile = reportsdir+'/Table3.3.csv'

    	try:
        	inFile = open(inputFile,'r')
    	except IOError:
        	print "[ERROR] Cound't open file: " + inputFile + ". Failed to generate Table 3.3.csv"
        	return None
    
    	try:
        	outFile = open(outputFile,'w')
    	except IOError:
        	print "[ERROR] Cound't write"+outputFile
        	return None
 
    	outFile.write('Tx Node, Shelf/Slot/Port, Equipment, Rx Node, Shelf/Slot/Port, Equipment, Power Delta (dB), Attenuator Loss, DCU Loss\n')  
    	yellowDHI = 0
    	for line in inFile:    
        	line = re.sub('[\n\r]','',line)    
        	flds = line.split(',')

        	if (is_number(flds[6])) and ((math.fabs(float(flds[6]))) > 1.0) and ((math.fabs(float(flds[6]))) <= 2.0):
            		if 'DCU' in flds[12]:
                		outFile.write(flds[0]+','+flds[4]+','+flds[3]+','+flds[8]+','+flds[11]+','+flds[10]+','+flds[6]+','+flds[12]+','+flds[13]+'\n')    
                		yellowDHI = yellowDHI + 1
    	if yellowDHI == 0:
        	outFile.write('N/A,N/A,N/A,N/A,N/A,N/A,N/A,N/A,N/A\n')
    	outFile.close()     
    	inFile.close()
    	return yellowDHI

def createTable42(reportsdir, processdir):
    
	inputFile = processdir+'/zRTRV-ALMTH.csv'
        pRTRVINV = processdir+'/zRTRV-INV.csv'
	pRTRVOCH = processdir+'/zRTRV-OCH.csv'
	outputFile = reportsdir+'/Table4.2.csv'
	pRTRVEQP = processdir+'/zRTRV-EQPT.csv'
	eqptFile = open(pRTRVEQP)
	eqpt = eqptFile.readlines()
	eqptFile.close()
	print eqpt[4][0] 	
	#outputFile = 'Table4.2.csv'
    	greenDHI = 0
    	yellowDHI = 0
    	redDHI = 0
    	try:
        	inFile = open(inputFile,'r')
    	except IOError:
        	print "[ERROR] Cound't open file: " + inputFile
        	return None
    
    	try:
        	outFile = open(outputFile,'w')
    	except IOError:
        	print "[ERROR] Cound't write"+outputFile
        	return None
    
	outFile.write('Node Name, Facility, Equipment, Pluggable, High, Low, Rx Power, Rx DHI\n')  

    	for line in inFile:    
        	line = re.sub('[\n\r]','',line)    
        	flds = line.split(',')
        
        	if is_number(flds[16]) and is_number(flds[17]):
            
            		try: 
                		inFile2 = open(pRTRVINV)
            		except IOError:
                		print "[ERROR] Couldn't open file:_pRTRV-INV.csv. Failed to generate Table 4.2.csv" 
            
           
            		for line2 in inFile2:
                		line2 = re.sub('[\n\r]','',line2)
                		flds2 = line2.split(',')
                		if flds[0] == flds2[0]:
                    			Node = flds[2].strip('ABCDEFGHIJKLMNOPQRSTUVWXYZ')
                    			Nodes = Node.split('-')
                    			slots = flds2[2].strip('ABCDEFGHIJKLMNOPQRSTUVWXYZ')
                    			slots = slots.split('-')
                    			if len(slots)>= 3:
                        			if Nodes[1] == slots[1] and Nodes[2]==slots[2]:
                            				if 'PPM' in flds2[2]:
                                				pluggable = 'Y'
                            
                            				else:
                                				pluggable = ''
                            				try:
                                				inFile3 = open(pRTRVOCH)
                            				except IOError:
                                				print "[ERROR} Couldn't open file:_pRTRV-OCH.csv. Failed to generate Table 4.2.csv"
                
                            				for line3 in inFile3:
                                				line3 = re.sub('[\n\r]','',line3)
                                				flds3 = line3.split(',')
                                				if flds3[0]==flds2[0] and flds3[2] == flds[2]:
                                    					checkWords = ['IS-NR','is-nr','Is-Nr','IS-Nr','Is-NR','unlocked-enabled']
                                    					isNR = flds3[8].split('|')
                                    					for i in isNR:
                                        					if i in checkWords:
                                            						if flds3[7] != '':
                                                						b = float(flds[17])
                                                						a = float(flds[16])
                                                						c= (math.fabs(b-a))/2
                                                						threshold75a = a + (c*0.25)
                                                						threshold75b = b - (c*0.25)
                                                						if ((float(flds3[7])) > c) and ((float(flds3[7])) <= threshold75b):
                                                    							DHI = 'Green'
                                                    							greenDHI = greenDHI+1
                                                						elif ((float(flds3[7])) >= threshold75a) and ((float(flds3[7])) < c):
                                                    							DHI = 'Yellow'
                                                    							yellowDHI = yellowDHI + 1
                                                						else:
                                                    							DHI = 'Red'   
                                                    							redDHI = redDHI + 1
                                                						outFile.write(str(flds[0])+','+str(flds[2])+','+str(flds2[3])+','+pluggable+','+str(flds[17])+','+str(flds[16])+','+str(flds3[7])+','+DHI+'\n')
                        
                          
            		inFile2.close()               
    	DHI = greenDHI + yellowDHI + redDHI
   	DHI = ((greenDHI*5) + (yellowDHI*3) + (redDHI*1))/DHI
    	return DHI

def createTable41(reportsdir, processdir):
    	
	inputFile = processdir+'/zRTRV-ALMTH.csv'
        outputFile = reportsdir+'/Table4.1.csv'
        pRTRVINV = processdir+'/zRTRV-INV.csv'
	pRTRVOCH = processdir+'/zRTRV-OCH.csv'
        pRTRVEQPT = processdir+'/zRTRV-EQPT.csv'

        getEQPT = csv.DictReader(open(pRTRVEQPT))
	Eqpt = {}
	for row in getEQPT:
		for column, value in row.iteritems():
			Eqpt.setdefault(column, []).append(value)
	#print Eqpt
	#sys.exit()	

    	greenDHI = 0
    	yellowDHI = 0
    	redDHI = 0
    	try:
        	inFile = open(inputFile,'r')
    	except IOError:
        	print "[ERROR] Cound't open file: " + inputFile
        	return None
    
    	try:
        	outFile = open(outputFile,'w')
    	except IOError:
        	print "[ERROR] Cound't write"+outputFile
        	return None
    	outFile.write('Node Name, Facility, Equipment, Pluggable, High, Low, Tx Power, Tx DHI\n')  

    	for line in inFile:    
        	line = re.sub('[\n\r]','',line)    
        	flds = line.split(',')
        
        	if is_number(flds[18]) and is_number(flds[19]):
            
            		try: 
                		inFile2 = open(pRTRVINV)
            		except IOError:
                		print "[ERROR] Couldn't open file:_zRTRV-INV.csv. Failed to generate Table 4.1.csv" 
            
           
            		for line2 in inFile2:
                		line2 = re.sub('[\n\r]','',line2)
                		flds2 = line2.split(',')
                		if flds[0] == flds2[0]:
                    			Node = flds[2].strip('ABCDEFGHIJKLMNOPQRSTUVWXYZ')
                    			Nodes = Node.split('-')
                    			slots = flds2[2].strip('ABCDEFGHIJKLMNOPQRSTUVWXYZ')
                    			slots = slots.split('-')
                    			if len(slots)>= 3:
                        			if Nodes[1] == slots[1] and Nodes[2]==slots[2]:
                            				if 'PPM' in flds2[2]:
                                				pluggable = 'Y'
                            
                            				else:
                                				pluggable = ''
                            				try:
                                				inFile3 = open(pRTRVOCH)
                            			
							except IOError:
                                				print "[ERROR} Couldn't open file:_pRTRV-OCH.csv. Failed to generate Table 4.1.csv"
                
                            				for line3 in inFile3:
                                				line3 = re.sub('[\n\r]','',line3)
                                				flds3 = line3.split(',')
                                				if flds3[0]==flds2[0] and flds3[2] == flds[2]:
                                    					checkWords = ['IS-NR','is-nr','Is-Nr','IS-Nr','Is-NR','unlocked-enabled']
                                    					isNR = flds3[8].split('|')
                                    					for i in isNR:
                                        					if i in checkWords:
                                            						if flds3[6] != '':
                                                						b = float(flds[19])
                                                						a = float(flds[18])
                                                						c= (math.fabs(b-a))/2
                                                						threshold75a = a + (c*0.25)
                                                						threshold75b = b - (c*0.25)
                                                						if ((float(flds3[6])) > c) and ((float(flds3[6])) <= threshold75b):
                                                    							DHI = 'Green'
                                                    							greenDHI = greenDHI+1
                                                						elif ((float(flds3[6])) >= threshold75a) and ((float(flds3[6])) < c):
                                                    							DHI = 'Yellow'
                                                    							yellowDHI = yellowDHI + 1
                                                						else:
                                                    							DHI = 'Red'   
                                                    							redDHI = redDHI + 1
                                                						outFile.write(str(flds[0])+','+str(flds[2])+','+str(flds2[3])+','+pluggable+','+str(flds[19])+','+str(flds[18])+','+str(flds3[6])+','+DHI+'\n')
                        
                          
        		inFile2.close()               
    	DHI = greenDHI + yellowDHI + redDHI
    	DHI = ((greenDHI*5) + (yellowDHI*3) + (redDHI*1))/DHI
    	return DHI

def createTable43(reportsdir, processdir):
    
 	inputFile = processdir+'/zRTRV-OCH.csv'
        outputFile = reportsdir+'/Table4.3.csv'


    	try:
        	inFile = open(inputFile,'r')
    	except IOError:
        	print "[ERROR] Cound't open file: " + inputFile +'. Failed to generate Table4.3.csv'
    	try:
        	outFile = open(outputFile,'w')
    	except IOError:
        	print "[ERROR] Cound't write Table4.3.csv"

    	outFile.write('Node Name, Facility, FEC Status, FEC Type\n')   
    	
	for line in inFile:
        	line = re.sub('[\n\r]','',line)    
        	flds = line.split(',')
        	if len(flds) >= 17 and flds[17] != '' and flds[0] != '#NODE':
            		if flds[17] == 'OFF':
                		FECStatus = 'OFF'
                		FECType = ''
            		else:
                		FECStatus = 'ON'
                		FECType = flds[17]
        
            		outFile.write(flds[0]+','+flds[2]+','+FECStatus+','+FECType+'\n')

    	return None

def createTable44(reportsdir, processdir):
    
    	inputFile = processdir+'/zRTRV-PM-ALL.csv'
        outputFile = reportsdir+'/Table4.4.csv'

	holdIP = []
    	#outputFile = 'Table4.4.csv'
    	greenDHI = 0
    	yellowDHI = 0
    	redDHI = 0 
    	try:
        	inFile = open(inputFile,'r')
    	except IOError:
        	print "[ERROR] Cound't open file: " + inputFile + ". Failed to generate Table 4.4.csv"
    	try:
        	outFile = open(outputFile,'w')
    	except IOError:
        	print "[ERROR] Cound't write Table4.4.csv"
   
    	for line in inFile:
        	line = re.sub('[\n\r]','',line)    
        	flds = line.split(',')
        	if flds[0] == '#NODE':
            		outFile.write('Node Name, Facility, Pre-FEC Bit Error Count, Uncorrected Word Count, DHI\n')
        	else:
			if flds[4] == 'BIT-EC':
                		if holdIP:
                    			outFile.write(holdIP[0]+','+holdIP[2]+','+holdIP[5]+','+UNCW+','+DHI+'\n') 
      
                    			holdIP = flds
                    			DHI = ''
                    			UNCW = '0'
            		elif (flds[4] == 'UNC-Words') and (holdIP[1] == flds[1]):
                		if flds[5] > 0:
                    			DHI =  'Red'
                    			redDHI = redDHI + 1
                		elif (holdIP[5] > 0 and flds[5] == 0):
                    			DHI = 'Yellow'
                    			yellowDHI = yellowDHI + 1
                		else:
                    			DHI = 'Green'
                    			greenDHI = greenDHI + 1
                		UNCW = flds[5]
            		elif holdIP:
                		if holdIP[1] != flds[1]:
                    			outFile.write(holdIP[0]+','+holdIP[2]+','+holdIP[5]+','+UNCW+','+DHI+'\n')  
                    			holdIP = []
    	DHI = greenDHI+yellowDHI+redDHI
    	if DHI != 0:
        	DHI = ((greenDHI*5)+(yellowDHI*3)+(redDHI))/DHI
    	return DHI    

def createTable45(reportsdir, processdir):
    
    	greenDHI = 0
    	yellowDHI = 0
    	redDHI = 0
    	holdIP = []
  	
	inputFile = processdir+'/zRTRV-OCH.csv'
        outputFile = reportsdir+'/Table4.5.csv'

    	#outputFile = 'Table4.5.csv'

    	try:
        	inFile = open(inputFile,'r')
    	except IOError:
        	print "[ERROR] Cound't open file: " + inputFile +". Failed to generate Table 4.5.csv"
    	try:
        	outFile = open(outputFile,'w')
    	except IOError:
        	print "[ERROR] Cound't write Table4.5.csv"
   
    	for line in inFile:
        	line = re.sub('[\n\r]','',line)    
        	flds = line.split(',')
        	if flds[2]=='AID':
            		outFile.write('#Node Name,Facility,WaveLength,Laser Bias%,DHI\n')    
        	elif len(flds) > 18:
            		if flds[18] and flds[18] != '0.0':
                		if float(flds[18]) > 90:
                    			DHI = 'RED'
                    			redDHI = redDHI + 1
                		elif float(flds[18]) < 75:
                    			DHI = 'GREEN'
                    			greenDHI = greenDHI + 1
                		else:
                    			DHI = 'YELLOW'
                    			yellowDHI = yellowDHI + 1
                		outFile.write(flds[0]+','+flds[2]+','+flds[4]+','+flds[18]+','+DHI+'\n')
      
    	DHI = redDHI + yellowDHI + greenDHI
    	DHI = (redDHI+yellowDHI*3+greenDHI*5)/DHI
    	return DHI

def createTable51(reportsdir, processdir):
    

 	inputFile = processdir+'/EQPT-SLOT.csv'
        outputFile = reportsdir+'/Table5.1.csv'

    	oldNodeName = 'X'
    
    	try:
        	inFile = open(inputFile,'r')
		#workbook = xlrd.open_workbook(inputFile)
    	except IOError:
        	print "[ERROR] Cound't open file: " + inputFile + ". Failed to generate Table 5.1.csv"
        	sys.exit()
    
    	try:
        	outFile = open(outputFile,'w')
    	except IOError:
        	print "[ERROR] Cound't write"+outputFile
        	sys.exit()
     
    	outFile.write('Node Name, Shelf/Slot, Equipment, Discrepancy\n')
	
	for line in inFile:
        	line = re.sub('[\n\r]','',line)
		row=line.split(',')
		
		if (row[10]) and (row[0] != '#Node'):
            		if row[6] or row[8]:
                		if row[4] and row[3]:
                    			columnB = str(int(row[3]))+'/'+str(int(row[4]))
                
                    			Disc = ''
                    			if row[10] == 'NOT in File1':
                        			Disc = 'Not in the CTP'
                    			elif row[10] == 'NOT in File2':
                        			Disc = 'Not in the NE'
                    			if row[6]:
                        			outFile.write(str(row[0])+','+columnB+','+str(row[6])+','+Disc+'\n')
                    			if row[8]:
                        			outFile.write(str(row[0])+','+columnB+','+str(row[8])+','+Disc+'\n')
    	return None

def createTable61(reportsdir, processdir):
    
  
    	
	inputFile = processdir+'/zRTRV-APC.csv'
        outputFile = reportsdir+'/Table6.1.csv'

	try:
        	inFile = open(inputFile,'r')
    	except IOError:
        	print "[ERROR] Cound't open file: " + inputFile +"Failed to generate Report 6.1.csv"
        	return None
    
    	try:
        	outFile = open(outputFile,'w')
    	except IOError:
        	print "[ERROR] Cound't write"+outputFile
        	return None
 
    	outFile.write('Node Name, Side, APC Enabled, Working Mode, DHI\n')  
    	greenDHI = 0
    	yellowDHI = 0
    
	for line in inFile:    
        	line = re.sub('[\n\r]','',line)    
        	flds = line.split(',')
        	if flds[1] != 'IP':
            		if flds[4] == 'Enabled' or flds[4] == 'WORKING':
                		greenDHI = greenDHI + 1
                		outFile.write(flds[0]+','+flds[2]+','+flds[3]+','+flds[4]+',Green\n')
            		else:
                		yellowDHI = yellowDHI + 1
                		outFile.write(flds[0]+','+flds[2]+','+flds[3]+','+flds[4]+',Green\n')

    	outFile.close()     
    	inFile.close()
    	DHI = (greenDHI+yellowDHI)
    	DHI = ((greenDHI*5)+(yellowDHI*3))/DHI
    	return DHI

def createTable62(reportsdir, processdir):
    
 	inputFile = processdir+'/zRTRV-ALS.csv'
        outputFile = reportsdir+'/Table6.2.csv'

    #outputFile = 'Table6.2.csv'
  
    	try:
        	inFile = open(inputFile,'r')
    	except IOError:
        	print "[ERROR] Cound't open file: " + inputFile +". Failed to generate Table 6.2.csv"
        	return None
    
    	try:
        	outFile = open(outputFile,'w')
    	except IOError:
        	print "[ERROR] Cound't write"+outputFile
        	return None
 
    	outFile.write('Node Name, Facility, ALS Mode, DHI\n')  
    	greenDHI = 0
    	yellowDHI = 0
    	for line in inFile:    
        	line = re.sub('[\n\r]','',line)    
        	flds = line.split(',')
        	if (flds[1] != 'IP') and (flds[3] != 'OCH' and flds[3] != 'och' and flds[3] != 'Och'):
            		if flds[4] == 'Disabled' or flds[4] == 'DISABLED':
                		yellowDHI = yellowDHI + 1
                		outFile.write(flds[0]+','+flds[2]+','+flds[4]+',Yellow\n')
            		else:
                		greenDHI = greenDHI + 1
                		outFile.write(flds[0]+','+flds[2]+','+flds[4]+',Green\n')

    	outFile.close()     
    	inFile.close()
    	DHI = (greenDHI+yellowDHI)
    	DHI = ((greenDHI*5)+(yellowDHI*3))/DHI
    	return DHI
     
        
def main():
  
	#inputFile = 'Process/inputfiles.csv' 
    	reportsdir = os.getcwd()
	os.chdir("..")
	processdir = "2.Process"
    	if not os.path.exists(reportsdir):
		try:
	    		os.makedirs(reportsdir)
        	except:
	    		print "[Error} Could not create", reportsdir
    
    	outputFile = reportsdir+'/'+'Params.csv'
    
    	#try:
        #	inFile = open(inputFile,'r')
    	#except IOError:
        #	print "[ERROR] Cound not open file: " + inputFile
        
    	try:
        	outFile = open(outputFile,'w')
    	except IOError:
        	print "[ERROR] Cound't write"+outputFile
        	sys.exit()
        
    	CUSTOMERNAME = ''
    	NCENAME = ''
    	while not CUSTOMERNAME:
    		CUSTOMERNAME = raw_input('Please enter Customer Name (no ^*^<^>^"^|^/ ) :  ')
    	while not NCENAME:
    		NCENAME = raw_input('Enter NCE name to be used as the Author Name of the report : ')

    	outFile.write('<Customer>|'+CUSTOMERNAME+'\n')


    	#for line in inFile:
        #	line = re.sub('[\n\r]','',line)    
        #	filenames = line.split(',')
    
    	#for inputFile in filenames:
        #	if 'WDMANS_OTS-OCH_TH' in inputFile:
            	# Getting Node Count
        #    		inputFile = 'Process/'+inputFile
	nodeCount = getNodeCount(processdir)
          	# Generating Section 2 Tables  
                	#Table21out = reportsdir+'/'+'Table2.1.csv'
	Table21DHI = createTable21(reportsdir, processdir)
        Table21DHIFlag = setDHIFlag(Table21DHI)
            		
			#Table22out = reportsdir+'/'+'Table2.2.csv'
	Table22DHI = createTable22(reportsdir, processdir)
        Table22DHIFlag = setDHIFlag(Table22DHI)
            		
			#Table23out = reportsdir+'/'+'Table2.3.csv'
	Table23DHI = createTable23(reportsdir, processdir)
        Table23DHIFlag = setDHIFlag(Table23DHI)
			
			#Table24out = reportsdir+'/'+'Table2.4.csv'
        Table24DHI = createTable24(reportsdir, processdir)
        Table24DHIFlag = setDHIFlag(Table24DHI)
            		
	Figure21 = createFig21(reportsdir, processdir)
        section2DHI = (Table21DHI+Table22DHI+Table23DHI+Table24DHI)/4
        section2DHI = "{0:.2f}".format(section2DHI)
        section2DHIFlag = setDHIFlag(section2DHI)  
          	# Generating Section 3 Tables
        #	if '_Report_' in inputFile:
        Table31DHI = createTable31(reportsdir, processdir)
        Table32DHI = 1
        section3DHI = createTable32(reportsdir, processdir)
        Table32DHIFlag = setDHIFlag(Table32DHI)
        Table33DHI = createTable33(reportsdir, processdir)
            
            		#section3DHI = "{0:.2f}".format(section3DHI)
        section3DHIFlag = setDHIFlag(section3DHI)   
        
	#	if '_pRTRV-ALMTH' in inputFile:
        Table41DHI = createTable41(reportsdir, processdir)
        Table41DHIFlag = setDHIFlag(Table41DHI)
        Table42DHI = createTable42(reportsdir, processdir)
        Table42DHIFlag = setDHIFlag(Table42DHI)
        
	#	if '_pRTRV-OCH' in inputFile:
        Table43DHI = createTable43(reportsdir, processdir)
            		#Table43DHIFlag = setDHIFlag(Table43DHI)
        Table45DHI = createTable45(reportsdir, processdir)
        Table45DHIFlag = setDHIFlag(Table45DHI)
        
	#	if '_pRTRV-PM-ALL' in inputFile:
        Table44DHI = createTable44(reportsdir, processdir)
        Table44DHIFlag = setDHIFlag(Table44DHI)
            
   	#	if '_EQPT-SLOT' in inputFile:    
        Table51DHI = createTable51(reportsdir, processdir)
        
        #	if '_pRTRV-APC' in inputFile:
        Table61DHI = createTable61(reportsdir, processdir)
        Table61DHIFlag = setDHIFlag(Table61DHI)
        
        #	if '_pRTRV-ALS' in inputFile:
        Table62DHI = createTable62(reportsdir, processdir)
        Table62DHIFlag = setDHIFlag(Table62DHI)
       
    	section4DHI = (Table41DHI + Table42DHI + Table45DHI + Table44DHI)/4
    	#section4DHI = "{0:.2f}".format(section4DHI)
    	section4DHIFlag = setDHIFlag(section4DHI)
    	section6DHI = (Table61DHI + Table62DHI)/2
    	#section6DHI = "{0:.2f}".format(section6DHI)
    	section6DHIFlag = setDHIFlag(section6DHI)
    	netDHI = (float(section2DHI)+float(section3DHI)+float(section4DHI)+float(section6DHI))/4
    	netDHIFlag = setDHIFlag(netDHI)     
    	outFile.write('<netNodeCount>|'+str(nodeCount)+'\n')
    	outFile.write('<netReachableNodeCount>|'+str(nodeCount)+'\n')
    	outFile.write('<netCollectedNodeCount>|'+str(nodeCount)+'\n')
    	section2DHI = "{0:.2f}".format(float(section2DHI))
    	outFile.write('<section2_DHI>|'+section2DHIFlag+'|'+str(section2DHI)+'\n')
    	section3DHI = "{0:.2f}".format(float(section3DHI))
    	outFile.write('<section3_DHI>|'+section3DHIFlag+'|'+str(section3DHI)+'\n')
    	section4DHI = "{0:.2f}".format(float(section4DHI))
    	outFile.write('<section4_DHI>|'+section4DHIFlag+'|'+str(section4DHI)+'\n')
    	section6DHI = "{0:.2f}".format(float(section6DHI))
    	outFile.write('<section6_DHI>|'+section6DHIFlag+'|'+str(section6DHI)+'\n')
    	netDHI = "{0:.2f}".format(float(netDHI))
    	outFile.write('<netDHI_flag>|'+netDHIFlag+'|'+str(netDHI)+'\n')
    
    	try:
        	workbook = xlrd.open_workbook(reportsdir+'/dwdm_proj.xlsx')
   
    	except IOError:
        	print "[ERROR] Cound't open file: " + inputFile
        	return None
    	worksheet = workbook.sheet_by_name('Recommend')
    
	if Table21DHIFlag == 'GREEN':
        	section21Recommend = worksheet.cell_value(13, 2)
    	elif Table21DHIFlag == 'YELLOW':
        	section21Recommend = worksheet.cell_value(13, 3)
    	elif Table21DHIFlag == 'RED':
        	section21Recommend = worksheet.cell_value(13, 4)
    
    	outFile.write('<table2.1Recommend>|'+section21Recommend+'\n')
    
    	if Table22DHIFlag == 'GREEN':
        	section22Recommend = worksheet.cell_value(14, 2)
    	elif Table22DHIFlag == 'YELLOW':
        	section22Recommend = worksheet.cell_value(14, 3)
    	elif Table22DHIFlag == 'RED':
        	section22Recommend = worksheet.cell_value(14, 4)
    	outFile.write('<table2.2Recommend>|'+section22Recommend+'\n')
    
    	if Table23DHIFlag == 'GREEN':
        	section23Recommend = worksheet.cell_value(15, 2)
    	elif Table23DHIFlag == 'YELLOW':
        	section23Recommend = worksheet.cell_value(15, 3)
    	elif Table23DHIFlag == 'RED':
        	section23Recommend = worksheet.cell_value(15, 4)
    	outFile.write('<table2.3Recommend>|'+section23Recommend+'\n')
    
    	if Table24DHIFlag == 'GREEN':
        	section24Recommend = worksheet.cell_value(16, 2)
    	elif Table24DHIFlag == 'YELLOW':
        	section24Recommend = worksheet.cell_value(16, 3)
    	elif Table24DHIFlag == 'RED':
        	section24Recommend = worksheet.cell_value(16, 4)
    	outFile.write('<table2.4Recommend>|'+section24Recommend+'\n')
    
    	section32Recommend =  worksheet.cell_value(17, 4)
    	outFile.write('<table3.2Recommend>|'+section32Recommend+'\n')
    
    	section33Recommend = worksheet.cell_value(18, 3)
    	outFile.write('<table3.3Recommend>|'+section33Recommend+'\n')
    
    	if Table41DHIFlag == 'GREEN':
        	section41Recommend = worksheet.cell_value(19, 2)
    	elif Table41DHIFlag == 'YELLOW':
        	section41Recommend = worksheet.cell_value(19, 3)
    	elif Table41DHIFlag == 'RED':
        	section41Recommend = worksheet.cell_value(19, 4)
    
    	outFile.write('<table4.1Recommend>|'+section41Recommend+'\n')
    
    	if Table42DHIFlag == 'GREEN':
        	section42Recommend = worksheet.cell_value(20, 2)
    	elif Table42DHIFlag == 'YELLOW':
        	section42Recommend = worksheet.cell_value(20, 3)
    	elif Table42DHIFlag == 'RED':
        	section42Recommend = worksheet.cell_value(20, 4)
    	outFile.write('<table4.2Recommend>|'+section42Recommend+'\n')
    
    	if Table44DHIFlag == 'GREEN':
        	section44Recommend = worksheet.cell_value(21, 2)
    	elif Table44DHIFlag == 'YELLOW':
        	section44Recommend = worksheet.cell_value(21, 3)
    	elif Table44DHIFlag == 'RED':
        	section44Recommend = worksheet.cell_value(21, 4)
    	outFile.write('<table4.4Recommend>|'+section44Recommend+'\n')
    
    	if Table45DHIFlag == 'GREEN':
        	section45Recommend = worksheet.cell_value(22, 2)
    	elif Table45DHIFlag == 'YELLOW':
        	section45Recommend = worksheet.cell_value(22, 3)
    	elif Table45DHIFlag == 'RED':
        	section45Recommend = worksheet.cell_value(22, 4)
    	outFile.write('<table4.5Recommend>|'+section45Recommend+'\n')
    
    
    	if Table61DHIFlag == 'GREEN':
        	section61Recommend = worksheet.cell_value(23, 2)
    	elif Table61DHIFlag == 'YELLOW':
        	section61Recommend = worksheet.cell_value(23, 3)

    	outFile.write('<table6.1Recommend>|'+section61Recommend+'\n')
    
    	if Table62DHIFlag == 'GREEN':
        	section62Recommend = worksheet.cell_value(24, 2)
    	elif Table62DHIFlag == 'YELLOW':
        	section62Recommend = worksheet.cell_value(24, 3)
    	outFile.write('<table6.2Recommend>|'+section62Recommend+'\n')
    
    	if netDHIFlag == 'GREEN':
        	netRecommend = worksheet.cell_value(12, 2)
    	elif netDHIFlag == 'YELLOW':
        	netRecommend = worksheet.cell_value(12, 3)
    	elif netDHIFlag == 'RED':
        	netRecommend = worksheet.cell_value(12, 4)

    	outFile.write('<netRecommend>|'+netRecommend+'\n')
    	outFile.write('<authorName>|'+NCENAME+'\n')
    	outFile.write('<reportDate>|'+(strftime("%B")+' '+(strftime("%d"))+','+(strftime("%Y"))+'\n'))
    
    
if __name__ == '__main__':
    	main()
