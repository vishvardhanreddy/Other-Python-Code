'''
Created on Jul 3, 2014

@author: vvaka
'''

import re, sys, csv
#================================================================================

def is_number(s):
  try:
    float(s)
    return True
  except ValueError:
    return False


def main():
    
  if len(sys.argv) != 3:
    sys.exit('Usage: Tab31DCUTypeVsLoss.py infile outfile ')
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

  
  outFile.write('DCU Type, Loss (dB)\n')  
  count = 0
  for line in inFile:    
    line = re.sub('[\n\r]','',line)    
    flds = line.split(',')
    if is_number(flds[13]):
        if 'DCU' in (flds[12]):
            outFile.write(flds[12]+','+flds[13]+'\n')    
      
  outFile.close()     
  inFile.close()
  
#================================================================================
if __name__ == "__main__":
  main()
      