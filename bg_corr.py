##############################################################################
## bg_corr.py (Python 2.7)
## Function: Does a linear BG removal
## Author: Seah Zong Long
## Version: 1.0
## Last modified: 29/11/2019

## Changelog: 


## Instructions:
## 1. Running this prompts a folder select. All appropriate (processed) XPS text files will be activated.
## 2. On plotted graph, select 2 points that define a linear bg and confirm selection.
## 3. A background removal is done with the outputs as a new _BGC graph and text file.
##############################################################################

import os
import matplotlib.pyplot as plt
import pandas as pd
import Tkinter
import tkFileDialog

def process_file(j):        # opens file, plots graph, and calls onclick 
    global f, fig, graph_2, cid, xcoords, ycoords, x, y
    x = []
    y = []
    xcoords = []
    ycoords = []

    filename = wlist[j]


    f=open('%s.txt' %filename)      ## opens text file as f
    lines = f.readlines()       ## reads all lines
    for i in range(len(lines)):
        if i > 0:
            x.append(float(lines[i].split()[0]))
            y.append(float(lines[i].split()[1]))
    

    fig = plt.figure()
    ax = fig.add_subplot(111)
    ax.plot(x,y)
    graph_2, = ax.plot([], marker='o')
    plt.text(0.7, 0.9, filename, size = 30,transform = ax.transAxes)
    
    cid = fig.canvas.mpl_connect('button_press_event', onclick)
    fig.canvas.mpl_connect('figure_leave_event', leave_figure)
    f.close()
    return


def leave_figure(event):        #Resets coords upon figure leave
    global xcoords, ycoords
    xcoords = []
    ycoords = []  
    fig.canvas.draw()


def onclick(event):
    global xcoords, ycoords, m, c, j, f, wlist, x, y
 
    xcoords.append(event.xdata)
    ycoords.append(event.ydata)
    graph_2.set_xdata(xcoords)
    graph_2.set_ydata(ycoords)    
    fig.canvas.draw()           #Constant refresh
    

#   after selected 2 coords, a line will be drawn. If accept, click again, otherwise leave figure to reset
    if len(xcoords) == 2:                
        m = (ycoords[1]-ycoords[0])/(xcoords[1]-xcoords[0])
        c = ycoords[0] - m*xcoords[0]

#   success condition: 3 clicks

    if len(xcoords) == 3:  
        writer = pd.ExcelWriter('processed_BGC/%s _BGC.xlsx' %str(wlist[j]), engine='xlsxwriter')
        y_bg = [(m*x_entry + c) for x_entry in x]
        y_bgc = [y_i - y_bg_i for y_i,y_bg_i in zip(y,y_bg)]
        bgc_data = pd.DataFrame(zip(x,y_bgc))
        bgc_data.to_excel(writer)
        writer.save()
        
        fig.canvas.mpl_disconnect(cid)
        fig.clf()
        plt.close()
        
        ax = bgc_data[1].plot()
        ax.set_xlabel("Binding energy (eV)")
        ax.set_ylabel("Counts")
        plt.savefig('processed_BGC/%s _BGC.jpg' %str(wlist[j]))
        plt.close()
        
        j += 1
        try:
            process_file(j)     #Recursive call until IndexError - selects the next one in the folder
        except IndexError:
            return


############ File Search ############
root = Tkinter.Tk()
root.withdraw() #use to hide tkinter window
root.wm_attributes('-topmost', 1) # Forces askdirectory on top

currdir = os.getcwd()
tempdir = tkFileDialog.askdirectory(parent=root, initialdir='C:\\Users\\E0004621\\Desktop\\ONDL Computer sync\\Papers\\Data\\4. Energy level alignment\\XPS', title='Please select the data folder') #select directory for data
#tempdir = 'C:\\Users\\E0004621\\Desktop\\ONDL Computer sync\\Papers\\Data\\4. Energy level alignment\\XPS\\Trial 2' #Debugging use
os.chdir(tempdir)

filelist = os.listdir(os.getcwd())  # working dir

if os.path.exists('processed_BGC')!=True: # make file
    os.mkdir('processed_BGC')

wlist=[]
for i in xrange(len(filelist)):
    ext = os.path.splitext(filelist[i])[1]
    filename = os.path.splitext(filelist[i])[0]
    if ext  in ('.txt','.TXT') and '(' in filename:     #only individual processed XPS files selected
        if 'WIDE' not in filename and 'wide' not in filename:
            wlist.append(filename)
        
print wlist        

############ Run recursive process  ############
j = 0
process_file(j)
    