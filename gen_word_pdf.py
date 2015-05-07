from oodocx.oodocx import *
from util import *

longcontents = [ "Table5.7.csv", "Table5.8.csv"]

def getTableName(str):
    tableName = os.path.splitext(str)[0]
    return '<'+tableName.lower()+'>'

def stripList(list):
    for i, item in enumerate(list):
        for j in ['\xa8C', '\x96', '\x08', '\r']:
            rloc = list[i].find(j)
            if rloc > 0:
                list[i] = item[:rloc]

    return list


def handleInputFile(filename,doc,reptDir):
    reportDir = reptDir
    table_w = None
    nhi_list = ["<hw_nhi>","<sw_nhi>","<al_nhi>","<pa_nhi>","<dcn_nhi>","<sr_nhi>","<se_nhi>","<cpt_nhi>"]

    print filename
    if "Table0.1" in filename:
        with open(filename) as csv:
            idx = 0
            for line in csv:
                list = [x.strip() for x in line.split(',')]
                list = stripList(list)
                if idx :
                    spara = doc.search (nhi_list[idx-1], result_type = 'run')
                    modify_font(spara, bold=True, ccolor=list[1].lower())
                    doc.replace(nhi_list[idx-1], list[1])

                idx = idx + 1
        return

    if filename in longcontents:
        table_w = [1278, 1800, 2520, 3644]

    body = doc.body
    table_name = getTableName(filename)
    table_paragraph = doc.search (table_name, result_type='paragraph')
    table_pos = body.index(table_paragraph)
    doc.replace(table_name,'')

    filename = path_join(reportDir, filename)
    contentList = []


    try:
        with open(filename) as csv:
            for line in csv:
                list = [x.strip() for x in line.split(',')]
                list = stripList(list)
                contentList.append(list)
    except IOError:
        print("handleInputFile Fail : IOError : %s not exists"%filename)
        return

    tbl = table(contentList, True, table_w, 'dxa', 0, 'auto', {'all':{"val":"thick"},}, _doc=doc )
    doc.body.insert(table_pos, tbl)


def handleInputList(infiles, doc, reptDir):
    filenames = [x.strip() for x in infiles.split(' ')]
    for filename in filenames:
        handleInputFile(filename,doc,reptDir)

def modify_bold(s_pram, r_pram,doc):
   spara = doc.search (s_pram, result_type = 'run')
   if r_pram in ["GREEN", "RED", "YELLOW", "GRAY"]:
        modify_font(spara, highlight=r_pram.lower(), bold=True)
   else:
        modify_font(spara, bold=True)
   doc.replace(s_pram, r_pram)

def getSvrVersion(reptDir):
   return reptDir[reptDir.find("_v")+2:]

def generate_report(custName, nceName, hasCPT, in1, in2, in3, in4, in5, in6, in7, in8, in9,reptDir):
    reportDir = reptDir

    if hasCPT == "YES":
        doc = Docx(tempDoc)
    else:
        doc = Docx(noCPTDoc)

    handleInputList(in1,doc, reptDir)
    handleInputList(in2,doc, reptDir)
    handleInputList(in3,doc, reptDir)
    handleInputList(in4,doc, reptDir)
    handleInputList(in5,doc, reptDir)
    handleInputList(in6,doc, reptDir)
    handleInputList(in7,doc, reptDir)
    handleInputList(in8,doc, reptDir)

    if hasCPT == "YES":
        handleInputList(in9, doc, reptDir)

    section9NHI_flag = False
    section9NHI_Tag = "<section9_NHI>"
    try:
        with open(path_join(reptDir, 'Params.csv')) as csv:
            for line in csv:
                list = [x.strip() for x in line.split('|')]
                if "netNHI" in list[0] :
                    modify_bold(list[0]+"1", list[1], doc)
                    modify_bold(list[0]+"2", list[1], doc)
	        elif "NHI" in list[0] :
                    modify_bold(list[0], list[1],doc)
                elif "Customer" in list[0]:
                    custName = custName.replace("_"," ")
                    modify_bold(list[0]+"1", custName, doc)
                    modify_bold(list[0]+"2", custName, doc)
                elif "authorName" in list[0]:
                    nceName = nceName.replace("_"," ")
                    modify_bold(list[0], nceName, doc)
                else:
                    if section9NHI_Tag in list[0] :
                        section9NHI_flag = True
                    if hasCPT != "YES" and "table9" in list[0]:
                        pass
                    else:
                        modify_bold(list[0]+"1", list[1], doc)
                        modify_bold(list[0]+"2", list[1], doc)
                        modify_bold(list[0], list[1], doc)

                svr_version = getSvrVersion(reptDir)
                if svr_version:
                    modify_bold("<SVR_Version>", svr_version, doc)
                else:
                    print("Error:svr version is missing")
    except IOError:
        print("generate_report() Fail : IOError : %s not exists"%path_join(reptDir, 'Params.csv'))
        return

    if hasCPT == "YES" and section9NHI_flag == False:
         spara = doc.search (section9NHI_Tag, result_type = 'run')
         modify_font(spara, bold=True)
         doc.replace(section9NHI_Tag, "N/A")

    date = str(datetime.datetime.now().date()).replace('-','')
    newDocName = "health_check_report_"+date+".docx"
    newDocFile = path_join(reportDir, newDocName)
    doc.save(newDocFile)

    try:
        os.system("unoconv -f pdf %s"%newDocName)
    except:
        pass
