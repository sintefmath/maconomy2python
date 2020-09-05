import csv
import argparse
import datetime
import numpy as np
import matplotlib.pyplot as plt 
from scipy.optimize import curve_fit

delimiter = ";"

def linear_func(x, a, b):
    return a + b * x

def getbillingprice(fn):
    data={}
    with open(fn) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=delimiter)
        line_count = 0
        for row in csv_reader:
            if line_count == 0:
                billing_index=row.index('Billing Price')
                date_index=row.index('Entry Date')
            else:
                if row[date_index] in data:
                    data[row[date_index]] += int(row[billing_index].replace(" ", "").replace(",00",""))
                else:
                    data[row[date_index]] = int(row[billing_index].replace(" ", "").replace(",00",""))
            line_count += 1
    return data

def getemployees(fn):
    data={}
    with open(fn) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=delimiter)
        line_count = 0
        for row in csv_reader:
            if line_count == 0:
                emplno_index=row.index('Empl. No.')
                emplname_index=row.index('Empl. Name')
            else:
                data[row[emplno_index]] = str(row[emplname_index])
            line_count += 1
    return data

def getemployeesbillingprice(employees_by_number, fn):
    data={}
    for key in employees_by_number:
        data[key]={}
    with open(fn) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=delimiter)
        line_count = 0
        for row in csv_reader:
            if line_count == 0:
                emplno_index=row.index('Empl. No.')
                billing_index=row.index('Billing Price')
                date_index=row.index('Entry Date')
            else:
                if row[date_index] in data[row[emplno_index]]:
                    data[row[emplno_index]][row[date_index]] += int(row[billing_index].replace(" ", "").replace(",00",""))
                else:
                    data[row[emplno_index]][row[date_index]] = int(row[billing_index].replace(" ", "").replace(",00",""))
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
    ### there might be two people with the exact name, we need to use the Empl. No.
    employees_by_number = getemployees(args.filename)
    billings_by_employees_by_day = getemployeesbillingprice(employees_by_number, args.filename)

    billings_by_employees_total = {}
    billings_by_employees_by_month = {}
    for employeenr, billings in billings_by_employees_by_day.items():
        billings_by_employees_total[employeenr] = 0
        billings_by_employees_by_month[employeenr] = np.zeros(12)
        for billingday, value in billings.items():
            billings_by_employees_total[employeenr] += value
            date = datetime.datetime.strptime(billingday, "%d.%m.%Y")
            billings_by_employees_by_month[employeenr][date.month-1] += value

    billings_by_month = np.zeros(12)
    for billingday, billings in billings_by_day.items():
        date = datetime.datetime.strptime(billingday, "%d.%m.%Y")
        billings_by_month[date.month-1] += billings

    labels_month=('Jan','Feb','Mar','Apr','Mai','Jun','Jul','Aug','Sep','Oct','Nov','Des')

    plt.rcParams["figure.figsize"] = 12,8
    plt.rcParams["axes.titlesize"] = 24
    plt.rcParams["axes.labelsize"] = 20
    plt.rcParams["lines.linewidth"] = 3
    plt.rcParams["lines.markersize"] = 10
    plt.rcParams["xtick.labelsize"] = 16
    plt.rcParams["ytick.labelsize"] = 16
    plt.style.use('bmh')

#### actuals per month
    fig1, ax1 = plt.subplots()
    ax1.bar(np.linspace(1,12,12),billings_by_month/1000)
    ax1.set_ylabel('KNOK')
    ax1.set_xticks(np.arange(1,13))
    ax1.set_xticklabels(labels_month)
    if args.totalbudget:
        ax1.plot([1,12], [args.totalbudget/12, args.totalbudget/12],':k',label='mean')
        ax1.legend()
    ax1.set_title('Actuals per month')
    plt.savefig("actuals_per_month.png")

    ### per employee
    fig, (ax1,lax) = plt.subplots(ncols=2, gridspec_kw={"width_ratios":[4,1]})
    bot = np.zeros(12)
    for a, b in billings_by_employees_by_month.items():
        ax1.bar(np.linspace(1,12,12),b/1000, bottom=bot/1000, label=employees_by_number[a])
        bot+=b
    ax1.set_ylabel('KNOK')
    ax1.set_xticks(np.arange(1,13))
    ax1.set_xticklabels(labels_month)
    if args.totalbudget:
        ax1.plot([1,12], [args.totalbudget/12, args.totalbudget/12],':k',label='mean')
    h,l = ax1.get_legend_handles_labels()
    lax.legend(h,l, borderaxespad=0)
    lax.axis("off")
    ax1.set_title('Actuals per month per employee')
    plt.tight_layout()
    plt.savefig("actuals_per_month_per_employee.png")

    today = datetime.datetime.today()
    week=today.isocalendar()[1]
    month=today.month
    year=today.year

    today = datetime.datetime.today()
    week=today.isocalendar()[1]
    month=today.month
    year=today.year

    ### if the project is from a past year, set month to 12 and week to number of weeks last year
    if today.year > date.year:
        month = 12
        week = datetime.date(year-1, 12, 31).isocalendar()[1]



### actuals accumulated per month
    cumsum_billings_by_month=np.cumsum(billings_by_month)
    cumsum_billings_by_month[month:] = 0

    fig1, ax1 = plt.subplots()
    ax1.bar(np.linspace(1,12,12),cumsum_billings_by_month/1000)
    ax1.set_ylabel('KNOK')
    ax1.set_xticks(np.arange(1,13))
    ax1.set_xticklabels(labels_month)
    if args.totalbudget:
        ax1.plot([1,12], [args.totalbudget, args.totalbudget],'-k',label='total budget')

    if args.regressionON and month>2:
        ydata = cumsum_billings_by_month[:month]
        xdata = np.arange(1,month+1)
        popt, pcov = curve_fit(linear_func, xdata, ydata)
        ax1.plot(np.arange(1,13),linear_func(np.arange(1,13), popt[0], popt[1])/1000,':k',label='linear regression')
        ax1.plot(12,linear_func(12, popt[0], popt[1])/1000,'ko')
        ax1.legend()
    ax1.set_title('Actuals accumulated per month')
    plt.savefig("actuals_accumulated_per_month.png")

    ### per employee
    fig1, ax1 = plt.subplots()
    bot = np.zeros(12)
    for a, b in billings_by_employees_by_month.items():
        bc = np.cumsum(b)
        bc[month:] = 0
        ax1.bar(np.linspace(1,12,12),bc/1000, bottom=bot/1000, label=employees_by_number[a])
        bot+=bc
    ax1.set_ylabel('KNOK')
    ax1.set_xticks(np.arange(1,13))
    ax1.set_xticklabels(labels_month)
    if args.totalbudget:
        ax1.plot([1,12], [args.totalbudget, args.totalbudget],'-k',label='total budget')

    if args.regressionON and month>2:
        ydata = cumsum_billings_by_month[:month]
        xdata = np.arange(1,month+1)
        popt, pcov = curve_fit(linear_func, xdata, ydata)
        ax1.plot(np.arange(1,13),linear_func(np.arange(1,13), popt[0], popt[1])/1000,':k',label='linear regression')
        ax1.plot(12,linear_func(12, popt[0], popt[1])/1000,'ko')
        ax1.legend()
    ax1.set_title('Actuals accumulated per month per employee')
    plt.savefig("actuals_accumulated_per_month_per_employee.png")

    fig, (ax1,lax) = plt.subplots(ncols=2, gridspec_kw={"width_ratios":[4,1]})
    bot = np.zeros(12)
    for a, b in billings_by_employees_by_month.items():
        ax1.bar(np.linspace(1,12,12),b/1000, bottom=bot/1000, label=employees_by_number[a])
        bot+=b
    ax1.set_ylabel('KNOK')
    ax1.set_xticks(np.arange(1,13))
    ax1.set_xticklabels(labels_month)
    if args.totalbudget:
        ax1.plot([1,12], [args.totalbudget/12, args.totalbudget/12],':k',label='mean')
    h,l = ax1.get_legend_handles_labels()
    lax.legend(h,l, borderaxespad=0)
    lax.axis("off")
    ax1.set_title('Actuals per month per employee')
    plt.tight_layout()
    plt.savefig("actuals_per_month_per_employee.png")


#### actuals per week
    num_weeks = datetime.date(year, 12, 31).isocalendar()[1]
    billings_by_week = np.zeros(num_weeks)
    for billingday, billings in billings_by_day.items():
        date = datetime.datetime.strptime(billingday, "%d.%m.%Y")
        billings_by_week[date.isocalendar()[1]-1] += billings
    fig1, ax1 = plt.subplots()
    ax1.bar(np.linspace(1,num_weeks,num_weeks),billings_by_week/1000)
    ax1.set_ylabel('KNOK')
    if args.totalbudget:
        ax1.plot([1,num_weeks], [args.totalbudget/num_weeks, args.totalbudget/num_weeks],':k',label='mean')
        ax1.legend()
    ax1.set_title('Actuals per week')
    plt.savefig("actuals_per_week.png")


### actuals accumulated per week
    cumsum_billings_by_week=np.cumsum(billings_by_week)
    cumsum_billings_by_week[week:] = 0

    fig1, ax1 = plt.subplots()
    ax1.bar(np.linspace(1,num_weeks,num_weeks),cumsum_billings_by_week/1000)
    ax1.set_ylabel('KNOK')
    if args.totalbudget:
        ax1.plot([1,num_weeks], [args.totalbudget, args.totalbudget],'-k',label='total budget')

    if args.regressionON and week>2:
        ydata = cumsum_billings_by_week[:week]
        xdata = np.arange(1,week+1)
        popt, pcov = curve_fit(linear_func, xdata, ydata)
        ax1.plot(np.arange(1,num_weeks+1),linear_func(np.arange(1,num_weeks+1), popt[0], popt[1])/1000,':k',label='linear regression')
        ax1.plot(num_weeks,linear_func(num_weeks, popt[0], popt[1])/1000,'ko')
        ax1.legend()
    ax1.set_title('Actuals accumulated per week')
    plt.savefig("actuals_accumulated_per_week.png")

### pie charts
    employeenames = list(employees_by_number.values())
    pie_sizes = np.array(list(billings_by_employees_total.values()))
    usedbudget = np.sum(pie_sizes)

    fig1, ax1 = plt.subplots()
    ax1.pie(pie_sizes/usedbudget, labels=employeenames, autopct='%1.1f%%', shadow=True, startangle=90)
    ax1.axis('equal')
    ax1.set_title('Budget actuals')
    plt.savefig("pie1.png")

    if args.totalbudget:
        explode = np.append(np.zeros_like(pie_sizes), 0.1)
        pie_labels = employeenames.copy()
        pie_labels.append('remaining')
        pie_sizes = np.append(pie_sizes, args.totalbudget*1000-usedbudget)
        pie_sizes = pie_sizes/args.totalbudget*1000
        fig1, ax1 = plt.subplots()
        ax1.pie(pie_sizes, explode=explode, labels=pie_labels, autopct='%1.1f%%', shadow=True, startangle=90)
        ax1.axis('equal')
        ax1.set_title('Budget total')
        plt.savefig("pie2.png")

    ### print some stats
    ln = len(max(employeenames, key=len))
    tot={}
    for a,b in billings_by_employees_total.items():
        tot[a]=str(int(b/1000)).rjust(6, ' ')
    
    print("Billings [KNOK] (modulo round off errors):")
    tmp=str("Employee").ljust(ln, ' ')+" |"
    for i in range(0,12):
        tmp+=labels_month[i].rjust(4, ' ')
    tmp+="| total"
    print(tmp)

    tmp=str("-").ljust(ln, '-')+"--"
    for i in range(0,12):
        tmp+=str("-").rjust(4, '-')
    tmp+="-------"
    print(tmp)

    for a,b in billings_by_employees_by_month.items():
        tmp=employees_by_number[a].ljust(ln, ' ')+" |"
        for i in range(0,12):
            tmp+=str(int(b[i]/1000)).rjust(4, ' ')
        tmp+="|"+tot[a]
        print(tmp)

    tmp=str("-").ljust(ln, '-')+"--"
    for i in range(0,12):
        tmp+=str("-").rjust(4, '-')
    tmp+="-------"
    print(tmp)

    tmp=str(" ").ljust(ln, ' ')+" |"
    for i in range(0,12):
        tmp+=str(int(billings_by_month[i]/1000)).rjust(4, ' ')
    tmp+="|"+str(int(usedbudget/1000)).rjust(6,' ')
    print(tmp)
    tmp=str("-").ljust(ln, '-')+"--"
    for i in range(0,12):
        tmp+=str("-").rjust(4, '-')
    tmp+="-------"
    print(tmp)

    print("")
    print("Remaining:", args.totalbudget-int(usedbudget/1000), " KNOK")
    print("")
