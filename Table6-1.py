'''
Created on Jul 10, 2014

@author: vvaka
'''
import re, math, csv, sys

def is_number(s):
  try:
    float(s)
    return True
  except ValueError:
    return False

def main():
    
  if len(sys.argv) != 3:
    sys.exit('Usage: 71TabAPCStates infile outfile')
    return None
  
  inputFile = sys.argv[1]
  outputFile = sys.argv[2]  
  diff = []
  count = 0
  printCount = 0
  try:
    inFile = open(inputFile,'r')
  except IOError:
    print "[ERROR] Cound't open file: " + inputFile
    return None  
  try:
    outFile = open(outputFile,'w')
  except IOError:
    print "[ERROR] Cound't open file: " + outputFile
    return None
  outFile.write('Node Name,APC Enabled, Working Mode,DHI\n')
  for line in inFile:
    line = re.sub('[\n\r]','',line)    
    flds = line.split(',')
    if (flds[3] != 'Y' and flds[3] != 'APCENABLE') or (flds[4] != 'WORKING' and flds[4] != 'APCSTATE'):
      outFile.write(str(flds[0])+','+str(flds[3])+','+str(flds[4])+'YELLOW \n')
      printCount = printCount + 1
  if printCount == 0:
    outFile.write('N/A,N/A,N/A,N/A\n')
  inFile.close()  
  outFile.close()     

  
if __name__ == '__main__':
  main()