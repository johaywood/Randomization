import csv
import random

studyNumber = raw_input('Enter the study number: ')
startAnimalNum = int(input('Enter the starting animal number: '))
totalAnimals = int(input('Enter the total number of animals: '))
lastRow = int(totalAnimals + 1)
meanCell = int(totalAnimals + 3)
stDevCell = int(totalAnimals + 4)
count = 0
data = []
header = ["Animal #", "Weight (g)"]
    
while count < totalAnimals:
    data.append([startAnimalNum + count, ''])
    count += 1
    
random.shuffle(data)
data.append(['','',''])
data.append(['Mean', '=average($B$1:$B$%r)' % lastRow])
data.append(['StDev', '=stdev($B$1:$B$%r)' % lastRow])
data.append(['Range (20% of Mean)', '', '=$B$%r*0.8' % meanCell, 'to', '=$B$%r*1.2' % meanCell])
data.append(['Range (+- 3xStDev)', '', '=$B$%r-3*$B$%r' % (meanCell, stDevCell), 'to', '=$B$%r-3*$B$%r' % (meanCell, stDevCell)])


print(data)

with open('%s Randomization.csv' % studyNumber, 'wb') as fp:   
    a = csv.writer(fp, delimiter=',', quoting=csv.QUOTE_MINIMAL)
    a.writerow(header)
    a.writerows(data)