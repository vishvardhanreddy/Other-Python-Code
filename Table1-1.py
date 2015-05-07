'''
Created on Jul 31, 2014

@author: vvaka
'''
import re, sys, csv, xlrd, math

def main():
    
  if len(sys.argv) != 3:
    sys.exit('Usage: Table1-1.py infile outfile ')
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
  
  outFile.write('Hardware Platform,Software Version\n')  
  count = 0

  for line in inFile:    
    line = re.sub('[\n\r]','',line)    
    flds = line.split(',')
    if flds[16] != 'PLATFORM':
      outFile.write(flds[16]+','+flds[11]+'\n')
    
#================================================================================
if __name__ == "__main__":
  main()
    