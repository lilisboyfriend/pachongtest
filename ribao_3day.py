from openpyxl import load_workbook
import time
import csv

def getdate(i,day,m,month):
    inday = int(day) - i
    inday2 = int(day) - i + 1;
    if (inday < 1):
        remon = int(m) - 1
        inday = month[remon] + inday
    elif (inday < 10):
        inday = "0{}".format(inday)
    if (inday2 < 1):
        remon = int(m) - 1
        inday2 = month[remon] + inday2
    elif (inday2 < 10):
        inday2 = "0{}".format(inday2)
    return inday,inday2


def fhribao():
    m = time.strftime("%m", time.localtime())
    day = time.strftime("%d", time.localtime())
    wbmb = load_workbook("D:\日报\烽火并发日报模板.xlsx")
    wsmb = wbmb.active
    bz_list = [25, 13, 1]
    month = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]  # month list
    startday = int(day) - 2
    for i in range(1, 4):
        inday1, inday2 = getdate(i, day, m, month)
        file = "D:\日报\ss_2021-{}-{}.csv".format(m, inday1)
        print(file)
        csvF = open(file, 'r', encoding='utf-8')
        read = csv.reader(csvF)
        index = bz_list[i - 1]
        wsmb.cell(index, 1).value = "2021/{}/{}".format(m, inday2)
        index = index + 1
        for line in read:
            # print(line)
            k = 1
            for v in line:
                wsmb.cell(index, k).value = v
                k = k + 1
            index = index + 1
    wbmb.save("D:\日报\烽火并发日报2021-{}-{}到{}.xlsx".format(m, startday, day))

def zteribao():
    a=1



def main():
    fhribao()


if __name__ == '__main__':
   main()