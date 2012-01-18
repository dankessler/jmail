#Import what you need#

import imaplib
import getpass
import csv
import sys


#Get UserName
user=raw_input('GMail Username: ')
#Get Pass
passw = getpass.getpass()

#Get JoveFolder Name
imappath=raw_input('What''s your IMAP server address? ')
jovefld=raw_input('What is the name of your JoveReg folder in GMail? ')


curdir=sys.path[0]

csvout=curdir + '/JoveReg.csv'



#Open GMail Connections
gm=imaplib.IMAP4_SSL(imappath)
gm.login(user,passw)
gm.select(jovefld,readonly=True)



#Instantiate Body List
bodies=list()
textlist=list()
textlist.append('First Name')
textlist.append('Last Name')
textlist.append('Email')
textlist.append('Company')
textlist.append('Research Areas')
textlist.append('Department')
textlist.append('Title/Position')
textlist.append('User Name')
textlist.append('User ID')

#Define extraction function

out=csv.writer(file(csvout,'wb'),dialect='excel')

out.writerow(textlist)

#Iterate over emails
emltoti=len(gm.search(None,'ALL')[1][0].split())

emltot=gm.search(None,'ALL')[1][0].split()[emltoti-1]

for eml in gm.search(None,'ALL')[1][0].split():

   try:
      body=gm.fetch(eml,'(UID BODY[TEXT])')[1][0][1]
      body=body.replace('<BR/>','')
      body=body.replace('<blockquote>',' ')

      
      textbody=body.split()
      textlist=list()
      textlist.append(''.join(textbody[textbody.index('First')+2:textbody.index('Last')]))
      textlist.append(''.join(textbody[textbody.index('Last')+2:textbody.index('Email:')]))
      textlist.append(''.join(textbody[textbody.index('Email:')+1:textbody.index('Institution:')]))
      textlist.append(''.join(textbody[textbody.index('Institution:')+1:textbody.index('Areas:')-1]))
      textlist.append(''.join(textbody[textbody.index('Areas:')+1:textbody.index('Department:')]))
      textlist.append(''.join(textbody[textbody.index('Department:')+1:textbody.index('Position:')]) )      
      textlist.append(''.join(textbody[textbody.index('Position:')+1:textbody.index('Username:')]))
      textlist.append(''.join(textbody[textbody.index('Username:')+1:textbody.index('UserID:')]))
      textlist.append(''.join(textbody[textbody.index('UserID:')+1]))
      out.writerow(textlist)
      print('Currently on email #' + eml + ' of ' + emltot)


   except:
      continue

del(out)
