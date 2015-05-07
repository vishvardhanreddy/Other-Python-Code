'''
Created on Jul 3, 2014

@author: vvaka

'''

import re, sys, csv, math
#================================================================================

def is_number(s):
  try:
    float(s)
    return True
  except ValueError:
    return False


def main():
    
  if len(sys.argv) != 3:
    sys.exit('Usage: Table3-2.py infile outfile ')
    return None
  
  inputFile = sys.argv[1]
  outputFile = sys.argv[2]
  
  try:
    inFile = open(inputFile,'r')
  except IOError:
    print "[ERROR] Cound't open file: " + inputFile
    return None
    
  try:
    outFile = open(outputFile,'w')
  except IOError:
    print "[ERROR] Cound't write"+outputFile
    return False

  
  outFile.write('TX NODE,TX SHELF/SLOT/PORT, TX EQPT, RX NODE, RX SHELF/SLOT/PORT, RX EQPT, POWER DELTA, ATTENUATOR LOSS, DCU LOSS \n')  
  count = 0
  for line in inFile:    
    line = re.sub('[\n\r]','',line)    
    flds = line.split(',')
    if (is_number(flds[6])) and ((math.fabs(float(flds[6]))) > 2.0):
      if not flds[12]: 
        flds[12] = 'Info Not Available'         
      outFile.write(flds[0]+','+flds[4]+','+flds[3]+','+flds[8]+','+flds[11]+','+flds[10]+','+flds[6]+','+flds[12]+','+flds[13]+'\n')    
  outFile.close()     
  inFile.close()
  
    
#================================================================================
if __name__ == "__main__":
  main()