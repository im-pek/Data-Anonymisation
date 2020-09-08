import uuid
import pandas as pd
import numpy as np

df1 = pd.read_excel("filename.xlsx")#, sheet_name='Employees')

owner = df1['Owner: Full Name'].tolist()
originalmeetingnotes = df1['Notes'].tolist() #original meeting notes to anonymise
orgsattended = df1['Organisations Attended'].tolist()
contactsattended = df1['Contacts Attended'].tolist() #has a few entries in each cell, separated by semi-colons
createdby = df1['Created By: Full Name'].tolist()

orgsindex = list(np.arange(1, 1+len(orgsattended)))

clientids = []

for item in orgsindex:
    clientid = 'Client ' + str(item)
    clientids.append(clientid)
       
allcontacts= []
allcontactslist = []

for item in contactsattended:
    if not isinstance(item, float):
        splitted = item.split('; ')
        allcontactslist.append(splitted)
        for element in splitted:
            allcontacts.append(element)
    else:
        allcontactslist.append([])
           
finalallcontacts = []

for entity in allcontactslist:
    splits = []
    for x in entity:
        splits.extend(x.split())
    finalallcontacts.append(splits)

finalowner = []

for entity in owner:
        finalowner.append(entity.split())
       
finalorgsattended = []

for entity in orgsattended:
    finalorgsattended.append(entity.split())
   
finalcreatedby = []

for entity in createdby:
    finalcreatedby.append(entity.split())

for idx,thing in enumerate(allcontactslist):
    finalallcontacts[idx] == finalallcontacts[idx].extend(allcontactslist[idx])
      
for idx,thing in enumerate(owner):    
    finalowner[idx] == finalowner[idx].append(owner[idx])
   
for idx,thing in enumerate(orgsattended):
    finalorgsattended[idx] == finalorgsattended[idx].append(orgsattended[idx])
   
for idx,thing in enumerate(createdby):
    finalcreatedby[idx] == finalcreatedby[idx].append(createdby[idx])

anonymisedmeetingnotes = []

for ix,note in enumerate(originalmeetingnotes):
    splitnotes = note.split()
    if not isinstance(note, float):
        for listed in finalallcontacts[ix]:
            if not isinstance(listed, float):
                uuid1 = str(uuid.uuid4())
                uuidsliced1 = uuid1[:8]
                if listed in note:
                    note = note.replace(listed, uuidsliced1)
                for idx,thingy in enumerate(splitnotes):
                    if listed in splitnotes[idx] and listed in splitnotes[idx+1]:######
                        note = note.replace(listed, uuidsliced1)                
                    if listed in splitnotes[idx] and listed in splitnotes[idx+1] and listed in splitnotes[idx+2]:######
                        note = note.replace(listed, uuidsliced1)  
                    if listed in splitnotes[idx] and listed in splitnotes[idx+1] and listed in splitnotes[idx+2] and listed in splitnotes[idx+3]:######
                        note = note.replace(listed, uuidsliced1)                                                                    
        for element in finalowner[ix]:
            if not isinstance(element, float):
                uuid2 = str(uuid.uuid4())
                uuidsliced2 = uuid2[:8]
                if element in note:
                    note = note.replace(element, uuidsliced2)
                for idx,thingy in enumerate(splitnotes):                
                    if listed in splitnotes[idx] and listed in splitnotes[idx+1]:######
                        note = note.replace(element, uuidsliced2)                
                    if listed in splitnotes[idx] and listed in splitnotes[idx+1] and listed in splitnotes[idx+2]:######
                        note = note.replace(element, uuidsliced2)  
                    if listed in splitnotes[idx] and listed in splitnotes[idx+1] and listed in splitnotes[idx+2] and listed in splitnotes[idx+3]:######
                        note = note.replace(element, uuidsliced2)                                                                    
        for index,element in enumerate(finalorgsattended[ix]):
            if not isinstance(element, float):                        
                if element in note:
                    for inds,org in enumerate(orgsattended):
                        if element in org:            
                            note = note.replace(element, 'Client ' +  str(inds + 1))
                for idx,thingy in enumerate(splitnotes):
                    for inds,org in enumerate(orgsattended):
                        if element in org:                                
                            if element in splitnotes[idx] and element in splitnotes[idx+1]:######
                                note = note.replace(element, 'Client ' +  str(inds + 1))          
                            if element in splitnotes[idx] and element in splitnotes[idx+1] and element in splitnotes[idx+2]:######
                                note = note.replace(element, 'Client ' +  str(inds + 1))  
                            if element in splitnotes[idx] and element in splitnotes[idx+1] and element in splitnotes[idx+2] and listed in splitnotes[idx+3]:######
                                note = note.replace(element, 'Client ' +  str(inds + 1))                                                                                        
        for element in finalcreatedby[ix]:
            if not isinstance(element, float):
                uuid3 = str(uuid.uuid4())
                uuidsliced3 = uuid3[:8]
                if element in note:
                    note = note.replace(element, uuidsliced3)
                for idx,thingy in enumerate(splitnotes):
                    if element in splitnotes[idx] and listed in splitnotes[idx+1]:######
                        note = note.replace(element, uuidsliced3)                
                    if element in splitnotes[idx] and listed in splitnotes[idx+1] and listed in splitnotes[idx+2]:######
                        note = note.replace(element, uuidsliced3)  
                    if element in splitnotes[idx] and listed in splitnotes[idx+1] and listed in splitnotes[idx+2] and listed in splitnotes[idx+3]:######
                        note = note.replace(element, uuidsliced3)                                                                                                            

    note = note.replace('  ',' ')

    for item in clientids:
        note = note.replace(str(item) + ' ' + str(item) + ' ' + str(item) + ' ' + str(item) + ' ' + str(item), item)        
        note = note.replace(str(item) + ' ' + str(item) + ' ' + str(item) + ' ' + str(item), item)        
        note = note.replace(str(item) + ' ' + str(item) + ' ' + str(item), item)      
        note = note.replace(str(item) + ' ' + str(item), item)                

    #print (note)
    anonymisedmeetingnotes.append(note)
    #print (anonymisedmeetingnotes)

df = pd.DataFrame(anonymisedmeetingnotes, columns = ['Anonymised Meeting Notes'])
writer = pd.ExcelWriter('Anonymised Meeting Notes.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Anonymised Meeting Notes', index=False)
writer.save()