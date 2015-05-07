#===============================================================================
#==================== PROCESS EQPT DATA
#===============================================================================
def process_EQPT(array):
  EQPT = {}
  col_node_name = indexHEADER(array["HEADER"], "#NODE")
  col_node_ip   = indexHEADER(array["HEADER"], "IP")
  col_slot_id   = indexHEADER(array["HEADER"], "SLOT_ID")
  col_equipment = indexHEADER(array["HEADER"], "EQUIPMENT")
  col_cardname  = indexHEADER(array["HEADER"], "CARDNAME")
  col_cardmode  = indexHEADER(array["HEADER"], "CARDMODE")
  for i in array:
    if i != "HEADER":
      node      = str(array[i][col_node_name])
      ip        = str(array[i][col_node_ip])
      slot_id   = str(array[i][col_slot_id])
      equipment = str(array[i][col_equipment])
      cardname  = str(array[i][col_cardname])
      cardmode  = str(array[i][col_cardmode])
      if not (node,"MULTISHELF") in EQPT:
        EQPT[node,"MULTISHELF"] = False

      #===== PROCESS EQUIPMENT: SLOT
      if re.search(r'SLOT-\d+',slot_id):
        tmpEQPT = slot_id.replace('SLOT-','').split('-')
        if len(tmpEQPT) == 1:
          SHELF = '1'
          SLOT  = tmpEQPT[0]
        else:
          SHELF = tmpEQPT[0]
          SLOT  = tmpEQPT[1]

        if cardname and re.sub(r'[-_ ]','',cardname) != re.sub(r'[-_ ]','',equipment):
          EQPT[node,SHELF,SLOT] = equipment + " (" + cardname.replace(',',';') + ")"
        else:
          EQPT[node,SHELF,SLOT] = equipment

      #===== PROCESS EQUIPMENT: PUNIT
      elif re.search(r'PUNIT-\d+',slot_id):
        tmpEQPT = slot_id.replace('PUNIT-','').split('-')
        SHELF   = tmpEQPT[0]
        SLOT    = ""

        if cardname and re.sub(r'[-_ ]','',cardname) != re.sub(r'[-_ ]','',equipment):
          EQPT[node,SHELF,SLOT] = equipment + " (" + cardname.replace(',',';') + ")"
        else:
          EQPT[node,SHELF,SLOT] = equipment

      #===== PROCESS EQUIPMENT: SHELF
      elif re.search(r'SHELF-\d+',slot_id):
        EQPT[node,"MULTISHELF"] = True

      #===== PROCESS EQUIPMENT: PPM
      elif re.search(r'PPM-\d+',slot_id):
        pass

      else:
        print "...[RTRV-EQPT] not expected: " + node + " " + ip + " [" + slot_id + " - " + equipment + "]"

  #===== PROCESS PPM WITH SLOT EQUIPMENT (NEED TO COMPLETE SLOT EQUIPMENT FIRST)
  for i in array:
    if i != "HEADER":
      node      = str(array[i][col_node_name])
      ip        = str(array[i][col_node_ip])
      slot_id   = str(array[i][col_slot_id])
      equipment = str(array[i][col_equipment])
      cardname  = str(array[i][col_cardname])
      cardmode  = str(array[i][col_cardmode])

      #===== PROCESS EQUIPMENT: PPM
      if re.search(r'PPM-\d+',slot_id):
        tmpEQPT = slot_id.replace('PPM-','').split('-')
        if len(tmpEQPT) == 2:
          SHELF = '1'
          SLOT  = tmpEQPT[0] + '-' + tmpEQPT[1]
          RSLOT = tmpEQPT[0]
        else:
          SHELF = tmpEQPT[0]
          SLOT  = tmpEQPT[1] + '-' + tmpEQPT[2]
          RSLOT = tmpEQPT[1]

        if cardname:
          if re.sub(r'[-_ ]','',cardname) != re.sub(r'[-_ ]','',equipment):
            EQPT[node,SHELF,SLOT] = equipment + " (" + cardname.replace(',',';') + ")"
          else:
            EQPT[node,SHELF,SLOT] = equipment
        else:
          EQPT[node,SHELF,SLOT] = equipment

        if (node,SHELF,RSLOT) in EQPT:
          #EQPT[node,SHELF,SLOT] = EQPT[node,SHELF,SLOT] + " | CARD = " + EQPT[node,SHELF,RSLOT]
          EQPT[node,SHELF,SLOT] = EQPT[node,SHELF,SLOT] + " | " + EQPT[node,SHELF,RSLOT]

  return EQPT

#===============================================================================
#==================== FIND EQPT based on FAC_ID
#===============================================================================
def find_EQPT(NODE,TID,EQPT):
  SHELFSLOT = SHELF_SLOT_PORT(TID,EQPT[NODE,"MULTISHELF"])
  if (NODE,SHELFSLOT[1],SHELFSLOT[2]) in EQPT:
    return EQPT[NODE,SHELFSLOT[1],SHELFSLOT[2]]
  else:
    return ""

#===============================================================================
#==================== FIND UNIT/SHELF/SLOT/PORT_NAME based on FAC ID
#===============================================================================
def find_SHELF_SLOT_PORTNAME(NODE,TID,EQPT,PORTMAP):
  SHELFSLOT = SHELF_SLOT_PORT(TID,EQPT[NODE,"MULTISHELF"])
  SHELF = SHELFSLOT[1]
  SLOT  = SHELFSLOT[2]
  CARD  = find_EQPT(NODE,TID,EQPT)

  if not SLOT:
    SLOT = "-"

  if SHELFSLOT[4]:
    tmpPORT = SHELFSLOT[3] + "-" + SHELFSLOT[4]
  else:
    tmpPORT = SHELFSLOT[3]

  if CARD and SHELFSLOT[0] == "LINE":
    tmpEQPT = re.sub(r'[-_ ]', '', CARD)
    if (tmpEQPT, tmpPORT) in PORTMAP:
      PORTNAME = PORTMAP[tmpEQPT, tmpPORT]
    else:
      PORTNAME = tmpPORT
  else:
    PORTNAME = tmpPORT
  return (SHELF,SLOT,PORTNAME)

#===============================================================================
#==================== FIND UNIT/SHELF/SLOT/PORT_ID based on FAC ID
#===============================================================================
def SHELF_SLOT_PORT(TID,isMULTI):
  flds = TID.split('-')
  typeID  = ""
  shelfID = ""
  slotID  = ""
  portID  = ""
  txrxID  = ""
  if flds[0] == 'LINE':
    if len(flds) == 4:
      if isMULTI:
        typeID  = flds[0]
        shelfID = flds[1]
        slotID  = flds[2]
        portID  = flds[3]
        txrxID  = ""
      else:
        typeID  = flds[0]
        shelfID = "1"
        slotID  = flds[1]
        portID  = flds[2]
        txrxID  = flds[3]
    elif len(flds) == 5:
      typeID  = flds[0]
      shelfID = flds[1]
      slotID  = flds[2]
      portID  = flds[3]
      txrxID  = flds[4]
    else:
      print "...[SHELF_SLOT_PORT] LINE not expected: " + TID + " MULTISHELF = " + str(isMULTI)

  elif flds[0] == 'CHAN':
    if len(flds) == 3:
      typeID  = flds[0]
      shelfID = "1"
      slotID  = flds[1]
      portID  = flds[2]
      txrxID  = ""
    elif len(flds) == 4:
      if isMULTI:
        typeID  = flds[0]
        shelfID = flds[1]
        slotID  = flds[2]
        portID  = flds[3]
        txrxID  = ""
      else:
        typeID  = flds[0]
        shelfID = "1"
        slotID  = flds[1]
        portID  = flds[2]
        txrxID  = flds[3]
    elif len(flds) == 5:
      typeID  = flds[0]
      shelfID = flds[1]
      slotID  = flds[2]
      portID  = flds[3]
      txrxID  = flds[4]
    else:
      print "...[SHELF_SLOT_PORT] CHAN not expected: " + TID + " MULTISHELF = " + str(isMULTI)

  elif flds[0] == 'PLINE' or flds[0] == 'PCHAN':
    if len(flds) == 4:
      typeID  = flds[0]
      shelfID = flds[1]
      slotID  = ""
      portID  = flds[2]
      txrxID  = flds[3]
    else:
      print "...[SHELF_SLOT_PORT] PLINE/PCHAN not expected: " + TID + " MULTISHELF = " + str(isMULTI)

  elif flds[0] == 'FAC' or flds[0] == 'VFAC':
    if isMULTI and len(flds) == 5:
      typeID  = flds[0]
      shelfID = flds[1]
      slotID  = '-'.join(flds[2:-1])
      portID  = flds[-1]
      txrxID  = ""
    elif not isMULTI and len(flds) == 4:
      typeID  = flds[0]
      shelfID = "1"
      slotID  = '-'.join(flds[1:-1])
      portID  = flds[-1]
      txrxID  = ""
    else:
      print "...[SHELF_SLOT_PORT] FAC/VFAC not expected: " + TID + " MULTISHELF = " + str(isMULTI)

  else:
   print "...[SHELF_SLOT_PORT] TYPE not expected: " + TID + " MULTISHELF = " + str(isMULTI)
   return ("", "", "", "", "")
  return (typeID, shelfID, slotID, portID, txrxID)
