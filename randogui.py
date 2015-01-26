import tkinter as tk
import csv, random

class Randomization(tk.Frame):
    def __init__(self,master=None):
        tk.Frame.__init__(self, master)

        self.grid()
        
        button1 = tk.Button(self, text=u'New Rando', command=self.newStudyRando_window)
        button1.grid(column=0, row=0, padx=50, pady=50, sticky='NSEW')
        button2 = tk.Button(self, text=u'Outlier Check', command=self.outlierCheck_window)
        button2.grid(column=1, row=0, padx=50, pady=50, sticky='NSEW')
        
    def newStudyRando_window(self):
        self.grid_forget()
        self.newWindow1 = tk.Toplevel(self.master)
        self.app = newStudyRando(self.newWindow1)
        
    def outlierCheck_window(self):
        self.grid_forget()
        self.newWindow2 = tk.Toplevel(self.master)
        self.app = newOutlierCheck(self.newWindow2)
    
class newStudyRando(tk.Frame):
    def __init__(self,master=None):
        tk.Frame.__init__(self, master)
        
        self.grid()
        
        self.studyNumber = tk.StringVar()
        self.entry1 = tk.Entry(self, textvariable=self.studyNumber)
        self.entry1.grid(column=0, row=0, padx=8, pady=8, sticky='NSEW')
        self.studyNumber.set(u'Enter Study Number')
            
        self.startAnimalNum = tk.IntVar()
        self.entry2 = tk.Entry(self, textvariable=self.startAnimalNum)
        self.entry2.grid(column=0, row=1, padx=8, pady=8, sticky='NSEW')
        self.startAnimalNum.set(u'Enter Starting Animal Number')
        
        self.totalAnimals = tk.IntVar()
        self.entry3 = tk.Entry(self, textvariable=self.totalAnimals)
        self.entry3.grid(column=0, row=2, padx=8, pady=8, sticky='NSEW')
        self.totalAnimals.set(u'Enter Starting Animal Number')
        
        button = tk.Button(self, text=u'Randomize!', command=self.onRandoClick)
        button.grid(column=1, row=1, padx= 10, sticky='NSEW')
        
        self.grid_columnconfigure((0), weight=1, minsize=250)
        self.grid_columnconfigure((1), weight=1, minsize=125)
        
    def onRandoClick(self):

        self.VstudyNumber = self.studyNumber.get()
        self.VstartAnimalNum = self.startAnimalNum.get()
        self.VtotalAnimals = self.totalAnimals.get()
        
        count = 0
        data = []
        header = ["Animal #", "Weight (g)"]

        while count < self.totalAnimals.get():
            data.append([self.startAnimalNum.get() + count, ''])
            count += 1

        random.shuffle(data)
       
        with open('%s Randomization.csv' % self.studyNumber.get(), 'w', newline='') as fp:   
            a = csv.writer(fp, delimiter=',') #, quoting=csv.QUOTE_MINIMAL
            a.writerow(header)
            a.writerows(data)
            
class newOutlierCheck(tk.Frame):
    def __init__(self,master=None):
        tk.Frame.__init__(self, master)
        self.grid()

#        file_path = tk.FileDialog.askopenfilename()
        
        button3 = tk.Button(self, text=u'20% of Mean', command=self.newMean_window)
        button3.grid(column=0, row=0, padx=50, pady=50, sticky='NSEW')
        button4 = tk.Button(self, text=u'Mean +- 3 StDev', command=self.newStDev_window)
        button4.grid(column=1, row=0, padx=50, pady=50, sticky='NSEW')
        
    def newMean_window(self):     
        self.grid_forget()
        self.newWindow3 = tk.Toplevel(self.master)
        self.app = outlierCheck(self.newWindow3)
        
    def newStDev_window(self):    
        self.grid_forget()
        self.newWindow4 = tk.Toplevel(self.master)
        self.app = outlierCheck(self.newWindow4)
        
class outlierCheck(tk.Frame):
    def __init__(self,master=None):
        tk.Frame.__init__(self, master)
        self.grid()    
        
        rando = newStudyRando().onRandoClick()
        
        lastRow = rando.VtotalAnimals + 1
        meanCell = newStudyRando.onRandoClick().VtotalAnimals + 3
        stDevCell = newStudyRando.onRandoClick().VtotalAnimals + 4
        outlierCheckRows = []
        
        file_path = tkFileDialog.askopenfilename()
        
        outlierCheckRows.append(['','',''])
        outlierCheckRows.append(['Mean', '=average($B$1:$B$%r)' % lastRow])
        outlierCheckRows.append(['StDev', '=stdev($B$1:$B$%r)' % lastRow])
        outlierCheckRows.append(['Range (20% of Mean)', '', '=$B$%r*0.8' % meanCell, 'to', '=$B$%r*1.2' % meanCell])
        outlierCheckRows.append(['Range (+- 3xStDev)', '', '=$B$%r-3*$B$%r' % (meanCell, stDevCell), 'to', '=$B$%r-3*$B$%r' % (meanCell, stDevCell)])
        
        with open(filename, 'a', newline='') as fp:   
            a = csv.writer(fp, delimiter=',') #, quoting=csv.QUOTE_MINIMAL
            a.writerows(outlierCheckRows)
        
if __name__ == '__main__':
    root = tk.Tk()
    app = Randomization(master=root)
    app.mainloop()