import tkinter, csv, random

class Randomization_tk(tkinter.Frame):
    def __init__(self,master=None):
        tkinter.Frame.__init__(self, master)
        self.initialize()
        
    def initialize(self):
        self.grid()
        
        self.studyNumber = tkinter.StringVar()
        self.entry1 = tkinter.Entry(self, textvariable=self.studyNumber)
        self.entry1.grid(column=0, row=0, padx=8, pady=8, sticky='NSEW')
        self.studyNumber.set(u'Enter Study Number')
            
        self.startAnimalNum = tkinter.IntVar()
        self.entry2 = tkinter.Entry(self, textvariable=self.startAnimalNum)
        self.entry2.grid(column=0, row=1, padx=8, pady=8, sticky='NSEW')
        self.startAnimalNum.set(u'Enter Starting Animal Number')
        
        self.totalAnimals = tkinter.IntVar()
        self.entry3 = tkinter.Entry(self, textvariable=self.totalAnimals)
        self.entry3.grid(column=0, row=2, padx=8, pady=8, sticky='NSEW')
        self.totalAnimals.set(u'Enter Starting Animal Number')
        
        button = tkinter.Button(self, text=u'Randomize!', command=self.onRandoClick)
        button.grid(column=1, row=1, padx= 10, sticky='NSEW')
        
        self.grid_columnconfigure((0), weight=1, minsize=250)
        self.grid_columnconfigure((1), weight=1, minsize=125)
        
    def onRandoClick(self):

        lastRow = int(self.totalAnimals.get()) + 1
        meanCell = int(self.totalAnimals.get()) + 3
        stDevCell = int(self.totalAnimals.get()) + 4
        count = 0
        data = []
        header = ["Animal #", "Weight (g)"]

        while count < self.totalAnimals.get():
            data.append([self.startAnimalNum.get() + count, ''])
            count += 1

        random.shuffle(data)
        data.append(['','',''])
        data.append(['Mean', '=average($B$1:$B$%r)' % lastRow])
        data.append(['StDev', '=stdev($B$1:$B$%r)' % lastRow])
        data.append(['Range (20% of Mean)', '', '=$B$%r*0.8' % meanCell, 'to', '=$B$%r*1.2' % meanCell])
        data.append(['Range (+- 3xStDev)', '', '=$B$%r-3*$B$%r' % (meanCell, stDevCell), 'to', '=$B$%r-3*$B$%r' % (meanCell, stDevCell)])

        with open('%s Randomization.csv' % self.studyNumber.get(), 'w', newline='') as fp:   
            a = csv.writer(fp, delimiter=',') #, quoting=csv.QUOTE_MINIMAL
            a.writerow(header)
            a.writerows(data)
        
if __name__ == '__main__':
    root = tkinter.Tk()
    app = Randomization_tk(master=root)
    app.mainloop()