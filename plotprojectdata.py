import csv
import argparse
import datetime
import numpy as np
import matplotlib.pyplot as plt 
from scipy.optimize import curve_fit

def linear_func(x, a, b):
    return a + b * x

def getbillingprice(fn):
    data={}
    with open(fn) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        line_count = 0
        for row in csv_reader:
            if line_count > 0:
                if row[1] in data:
                    data[row[1]] += int(row[10].replace(" ", ""))
                else:
                    data[row[1]] = int(row[10].replace(" ", ""))
            line_count += 1
    return data

def getemployees(fn):
    data={}
    with open(fn) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        line_count = 0
        for row in csv_reader:
            if line_count > 0:
                data[row[18]] = str(row[19])
            line_count += 1
    return data

def getemployeesbillingprices(fn):
    data={}
    with open(fn) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        line_count = 0
        for row in csv_reader:
            if line_count > 0:
                if row[18] in data:
                    data[row[18]] += int(row[10].replace(" ", ""))
                else:
                    data[row[18]] = int(row[10].replace(" ", ""))
            line_count += 1
    return data

def str2bool(v):
    if isinstance(v, bool):
       return v
    if v.lower() in ('yes', 'true', 'True', 't', 'y', '1'):
        return True
    elif v.lower() in ('no', 'false', 'False', 'f', 'n', '0'):
        return False
    else:
        raise argparse.ArgumentTypeError('Boolean value expected.')


if __name__ == "__main__":

    parser = argparse.ArgumentParser(description='Parse and plot data')
    parser.add_argument('--filename', metavar='filename', required=True, type=str, help='name of cvs file')
    parser.add_argument('--totalbudget', metavar='totalbudget', required=False, type=int, help='total budget in KNOK')
    parser.add_argument('--regressionON', metavar='regressionON', type=str2bool, nargs='?', const=True, default=True, help='plot regression')
    args = parser.parse_args()
    billings_by_day = getbillingprice(args.filename)
    employees_by_number = getemployees(args.filename)
    billings_by_employees = getemployeesbillingprices(args.filename)
    tick_labels=('Jan','Feb','Mar','Apr','Mai','Jun','Jul','Aug','Sep','Oct','Nov','Des')

#### actuals per month
    billings_by_month = np.zeros(12)
    for datestr, billings in billings_by_day.items():
        date = datetime.datetime.strptime(datestr, "%d.%m.%Y")
        billings_by_month[date.month-1] += billings
    fig1, ax1 = plt.subplots()
    ax1.bar(np.linspace(1,12,12),billings_by_month/1000)
    ax1.set_ylabel('KNOK')
    ax1.set_xticks(np.arange(1,13))
    ax1.set_xticklabels(tick_labels)
    if args.totalbudget:
        ax1.plot([1,12], [args.totalbudget/12, args.totalbudget/12],':k',label='mean')
    ax1.set_title('Actuals per month')
    ax1.legend()
    plt.savefig("actuals_per_month.png")

### actuals accumulated
    billings_by_month = np.zeros(12)
    for datestr, billings in billings_by_day.items():
        date = datetime.datetime.strptime(datestr, "%d.%m.%Y")
        billings_by_month[date.month-1] += billings
    cumsum_billings_by_month=np.cumsum(billings_by_month)

    today = datetime.datetime.today()
    month=today.month
    ### if the project is from a past year, set month to 12
    if today.year > date.year:
        month=12

    cumsum_billings_by_month[month:] = 0
    fig1, ax1 = plt.subplots()
    ax1.bar(np.linspace(1,12,12),cumsum_billings_by_month/1000)
    ax1.set_ylabel('KNOK')
    ax1.set_xticks(np.arange(1,13))
    ax1.set_xticklabels(tick_labels)
    if args.totalbudget:
        ax1.plot([1,12], [args.totalbudget, args.totalbudget],'-k',label='total budget')

    if args.regressionON and month>2:
        ydata = cumsum_billings_by_month[:month]
        xdata = np.arange(1,month+1)
        popt, pcov = curve_fit(linear_func, xdata, ydata)
        ax1.plot(np.arange(1,13),linear_func(np.arange(1,13), popt[0], popt[1])/1000,':k',label='linear regression')
        ax1.plot(12,linear_func(12, popt[0], popt[1])/1000,'ko')
        ax1.legend()
    ax1.set_title('Actuals accumulated')
    plt.savefig("actuals_accumulated.png")

### pie charts
    ### the following assumes the same ordering in employees_by_number and billings_by_employees which should be true
    pie_labels = list(employees_by_number.values())
    pie_sizes = np.array(list(billings_by_employees.values()))
    usedbudget = np.sum(pie_sizes)

    fig1, ax1 = plt.subplots()
    ax1.pie(pie_sizes/usedbudget, labels=pie_labels, autopct='%1.1f%%', shadow=True, startangle=90)
    ax1.axis('equal')
    ax1.set_title('Budget actuals')
    plt.savefig("pie1.png")

    if args.totalbudget:
        explode = np.append(np.zeros_like(pie_sizes), 0.1)
        pie_labels.append('remaining')
        pie_sizes = np.append(pie_sizes, args.totalbudget*1000-usedbudget)
        pie_sizes = pie_sizes/args.totalbudget*1000
        fig1, ax1 = plt.subplots()
        ax1.pie(pie_sizes, explode=explode, labels=pie_labels, autopct='%1.1f%%', shadow=True, startangle=90)
        ax1.axis('equal')
        ax1.set_title('Budget total')
        plt.savefig("pie2.png")
    
