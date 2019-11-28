

############################################################
####                                                    ####
####            XPS Data Plotter(Plot and Excel)        ####
####                                                    ####
####    Written by Jun Kai, for his Lovely Sunshine~    ####
####    Version 2.7(17/6/13), now with text exporter in ####
####    background correction format for Kgraph users   ####
####    Base program time taken: ~4.5hrs (21/12/12)     ####
####    Requires installation of pyExcelerator,         ####
####    numpy, xlrd and matplotlib python packages      ####
####    ARIALN.ttf for Arial Narrow text and labels     ####
####    Removed need for Arial Narrow font              ####
####                                                    ####
############################################################

##############################################################################
## XPS Data Plotter (Python 2.7)
## Function: XPS Data Plotter
## Author: Tan Jun Kai, updated by Seah Zong Long
## Version: 2.7.1
## Last modified: 28/11/2019

## Changelog: Updated os.chdir, added ability to handle duplicates

## Instructions:
## Similar to pytest. Open and select file dir that contains text files
##############################################################################

from pyExcelerator import *
import xlrd
import os,sys
from numpy import *
import numpy as np
import matplotlib.pyplot as plt  
import Tkinter
import tkFileDialog


root = Tkinter.Tk()
root.withdraw() #use to hide tkinter window
root.wm_attributes('-topmost', 1) # Forces askdirectory on top

currdir = os.getcwd()
tempdir = tkFileDialog.askdirectory(parent=root, initialdir='C:\\Users\\E0004621\\Desktop\\ONDL Computer sync\\Papers\\Data\\4. Energy level alignment\\XPS', title='Please select the data folder') #select directory for data
#tempdir = 'C:\\Users\\E0004621\\Desktop\\ONDL Computer sync\\Papers\\Data\\4. Energy level alignment\\XPS\\Trial'
os.chdir(tempdir)

filelist = os.listdir(os.getcwd())  # working dir


if os.path.exists('processed')!=True:
    os.mkdir('processed')
    
wlist=[]
for i in xrange(len(filelist)):
    ext = os.path.splitext(filelist[i])[1]
    filename = os.path.splitext(filelist[i])[0]
    if ext  in ('.txt','.TXT'):
        wlist.append(filename)

for file in wlist:

    if os.path.exists('processed/%s' %file)!=True:
        os.mkdir('processed/%s' %file)

    w = Workbook()  ## creates workbook

    count = 0
    position = [0 for x in range (0,20)]
    
    f=open('%s.txt' %file)      ## opens text file as f
#    print 'Processing %s' %file
    lines = f.readlines()       ## reads all lines

##### DETERMINING START POSITONS #######

    for i in xrange(len(lines)):
        words = lines[i].split()
        if words[0] == 'Region':
            position[count] = i
            count += 1

##### SET LABELS AND FILLED POSITION ARRAY ####

    posit = [0 for x in range(0,count+1)]
    label = [0 for x in range(0,count)]
    chem_name = [0 for x in range(0,count)]
    shellname = [0 for x in range(0,count)]
    shift = [0 for x in range(0,count)]
    header = [0 for x in range(0,count*4)]
    maxlength = 0

    for l in xrange(count):
        posit[l] = position[l]
        words = lines[posit[l]+1].split()
        label[l] = words[12] + "(" + str(l) + ")"

    posit[count] = len(lines)
    for l in xrange(count):
        if maxlength<(posit[l+1]-posit[l]-5):
            maxlength=(posit[l+1]-posit[l]-5)
#    print maxlength
    databank = array([['emptyslot' for x in range(0,count*4)] for y in range(0,maxlength)])
    print posit
    print label

##### READ PARAMETERS #####

    step = [0 for x in range(0,count)]
    sweep = [0 for x in range(0,count)]
    dwell = [0 for x in range(0,count)]
    mode = [0 for x in range(0,count)]
    caecrr = [0 for x in range(0,count)]
    mag = [0 for x in range(0,count)]
    channels = [0 for x in range(0,count)]
    maxenergy = [0 for x in range(0,count)]
    maxcounts = [0 for x in range(0,count)]
    for i in xrange(count):
        details = lines[posit[i]+1].split()
        step[i] = details[5]
        sweep[i] = details[6]
        dwell[i] = details[7]
        mode[i] = details[8]
        caecrr[i] = details[9]
        mag[i] = details[10]
        channels[i] = details[13]

##### EXTRACT ENERGY AND COUNTS ######

    for j in range(0,count):
        
        energy = [0.0 for x in range (0,posit[j+1]-posit[j]-5)]
        counts = [0.0 for x in range (0,posit[j+1]-posit[j]-5)]
        
        for k in range(posit[j]+5,posit[j+1]):
            digit_split = lines[k].split()
            if digit_split[0]=='Layer':         ### Section for Layer issue
                dump_energy = [0.0 for x in range (0,k-posit[j]-5)]
                dump_counts = [0.0 for x in range (0,k-posit[j]-5)]
                for i in range(0,k-posit[j]-5):
                    dump_energy[i] = energy[i]
                    dump_counts[i] = counts[i]
                energy = dump_energy
                counts = dump_counts
                break                           ### end section
            energy[k-posit[j]-5] = float(digit_split[0])
            counts[k-posit[j]-5] = float(digit_split[1])
            if maxcounts[j] < counts[k-posit[j]-5]:
                maxcounts[j] = counts[k-posit[j]-5]
                maxenergy[j] = energy[k-posit[j]-5]

##### INPUT INTO HEADER AND DATABANK FOR OVERALL TXT #####

        header[0+j*4]=label[j]+" eV" 
        header[1+j*4]=label[j]
        header[2+j*4]=label[j]+" eVbg"
        header[3+j*4]=label[j]+" bg"
        for n in xrange(len(energy)):
            databank[n,0+j*4]=str(energy[n])
            databank[n,2+j*4]=str(energy[n])
            databank[n,1+j*4]=str(counts[n])
            databank[n,3+j*4]=str(counts[n])

##### WRITING INTO EXCEL SHEET #####

        ws = w.add_sheet(label[j])
        ws.write(0,0,'Energy(eV)')
        ws.write(0,1,'Counts')
        ws.write(0,3,'Start')
        ws.write(0,4,'End')
        ws.write(0,5,'Step')
        ws.write(0,6,'Sweep')
        ws.write(3,3,'Dwell')
        ws.write(3,4,'Mode')
        ws.write(3,5,'CAE/CRR')
        ws.write(3,6,'Mag')
        ws.write(3,7,'No. of Channels')
        ws.write(1,3,energy[0])
        ws.write(1,4,energy[len(energy)-1])
        ws.write(1,5,step[j])
        ws.write(1,6,sweep[j])
        ws.write(4,3,dwell[j])
        ws.write(4,4,mode[j])
        ws.write(4,5,caecrr[j])
        ws.write(4,6,mag[j])
        ws.write(4,7,channels[j])
        ws.write(6,3,'Peak:')
        ws.write(6,4,maxenergy[j])
        
        for n in xrange(len(energy)):
            ws.write(n+1,0,energy[n])
            ws.write(n+1,1,counts[n])

#### WRITING INTO TEXT FILE FOR KGRAPH USERS #####


        f2=open('processed/%s/%s.txt' %(file,label[j]),'w')
        f2. writelines("eV\tcounts\n")
        for i in xrange(len(energy)):
            f2.writelines("%g\t%g\n" %(energy[i], counts[i]))
        f2.close()

##### STYLISE LABELS #######


#        for i in range(0,len(label)):
        l = list(label[j])
        if len(l) == 2:
            chem_name[j] = label[j]
            shellname[j] = ''
        elif len(l) == 3:
            chem_name[j] = l[0]
            shellname[j] = l[1]+l[2]
        elif len(l) == 4 and (label[j] != 'Wide' and label[j] != 'wide' and label[j]!='WIDE'):
            chem_name[j] = l[0]+l[1]
            shellname[j] = l[2]+l[3]
            shift[j] = 0.02
##        elif label[j] == ('Wide' or 'wide' or 'WIDE'):
        else:
            chem_name[j] = label[j]
            shellname[j] = ''

##### PLOTTING OF SPECTRA #####               
            
        fig=plt.figure(figsize=(7,10),dpi=100)
        ax=fig.add_subplot(111)
#        ax.xaxis.set_minor_locator(minorLocator)
        plt.plot(energy,counts)
        plt.text(0.7,0.9,chem_name[j],size=30,transform = ax.transAxes)
        plt.text(0.765+shift[j],0.89,shellname[j], size=18, transform=ax.transAxes)
        plt.xlim(energy[0],energy[len(energy)-1])
        plt.xlabel('Energy(eV)')
        plt.ylabel('Counts')
        plt.grid(True)

#### SAVING PNG PICTURE #####       
        
#        plt.legend(loc='upper right', prop={"size":"larger"})
        plt.savefig('processed/%s/%s_%s.png' %(file,file,label[j]))
        plt.close()

##### SAVING EXCEL SHEET #####

    w.save('processed/%s/%s_p.xls' %(file,file))
    print '%s processed, excel saved' %file

##### WRITING OVERALL TXT FOR BACKGROUND CORRECTION PURPOSES #####
    
    f3=open('processed/%s/%s_overall.txt' %(file,file),'w')
    for i in xrange(len(header)):
        f3.writelines('%s\t' %header[i])
    f3.writelines('\n')
    for i in xrange(maxlength):
        for j in xrange(count*4):
            if databank[i,j]=="emptyslot":
                databank[i,j]=''
            f3.writelines('%s\t' %databank[i,j])
        f3.writelines('\n')
    f3.close()
    f.close()
