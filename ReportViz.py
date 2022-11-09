#Matthew Trinh
#7/17/2020

'''
Function:
    Creates line plots from reports

Notes:
    All parameteres are optional - see -h in terminal and comments above reportdates function
    Generalize to other parent companies
        Easy: ask for path input of parent company's markets
        Hard: ask for parent company and use os.walk or glob recursion to find folders
    Add error codes
'''

import os
import glob

from datetime import datetime, date, timedelta

from openpyxl import workbook, load_workbook
import pandas as pd

import matplotlib.pyplot as plt
from matplotlib import rcParams as mpl

import argparse


#Command line arguments
parser = argparse.ArgumentParser()
parser.add_argument("-p", "--parent", help="path to markets | default: GateHouse")
parser.add_argument("-m", "--markets", nargs='+', help="markets separated by spaces | default: all")
parser.add_argument("-s", "--start", help="starting date (YYYYMMDD)")
parser.add_argument("-e", "--end", help="ending date (YYYYMMDD) | default: today")
parser.add_argument("-w", "--weeks", help="number of weeks", type=int)
parser.add_argument("-o", "--output", help="output directory")
args = parser.parse_args()

#Working directory
if not(args.parent):
    os.chdir('D:/data/GateHouse')
else:
    os.chdir(args.parent)

#If no markets specified, do all
if not(args.markets):
    args.markets = os.listdir()


#Finds dates within range
#Optional:starting date, ending date, number of weeks
'''
Not super elegant, but has litty functionality
NS NE NW - Last 10 weeks from today
NS NE W - Last w weeks from today
NS E NW - Last 10 weeks from end_date
NS E W - Last w weeks from end_date
S NE NW - 10 weeks from start_date
S NE W - w weeks from start_date
S E NW - From start_date to end_date
S E W - From start_date to end_Date
'''
def reportdates(market,start_date=None,end_date=None, weeks=None):
    end = True
    if(not(end_date)):
        end = False
    if(not(weeks)):
        weeks=10

    if(not(end)):
        end_date = date.today()
    else:
        end_date = datetime.strptime(end_date,'%Y%m%d')

    if(not(start_date)):
        start_date = (end_date - timedelta(weeks = weeks)).strftime('%Y%m%d')
    elif(not(end)):
        end_date = datetime.strptime(start_date,'%Y%m%d') + timedelta(weeks = weeks)

    end_date = end_date.strftime('%Y%m%d')


    alldates = glob.glob('./%s/report/%s*report*.xlsx'%(market,market))
    alldates = [''.join(filter(str.isdigit,x)) for x in alldates]

    dates = [x for x in alldates if x >= start_date and x <= end_date]
    dates = list(set(dates))
    dates.sort()

    return dates

for market in args.markets:

    if(os.path.exists('./%s/report' % (market))):

        df_main = pd.DataFrame()
        dates = reportdates(market,args.start,args.end,args.weeks)

        if(len(dates) == 0):
            continue

        for dt in dates:
            try:
                path = glob.glob('./%s/report/%s*report*%s*.xlsx' % (market,market,dt))

                #Reads 'Summary All Expired' sheet from report. Converts into pandas dataframe
                #Faster than pd.read_excel(path[0],sheet_name='Summary All Expired',engine='openpyxl')
                report = load_workbook(filename = path[0], data_only = True, read_only = True)
                ws = report['Summary All Expired']
                df = pd.DataFrame(ws.values)

                #Gets relevant rows and columns
                df = df.iloc[12:25,[0,2]]

                #Formatting
                df.columns=['metric','percent']
                df = df.set_index('metric')
                df['percent'] = df['percent']*100

                #Append to main dataframe
                df_main = pd.concat([df_main,df])
            except:
                print('Error for %s %s' % (market,dt))

        if(len(df) == 0):
            print('Visualization error for %s' % market)
            continue

        gb = df_main.groupby('metric')

        revert = gb.get_group('Reverts')['percent'].to_list()
        r2b = gb.get_group('Revert to Price Below Original')['percent'].to_list()
        r2o = gb.get_group('Revert to Original')['percent'].to_list()
        r2a = gb.get_group('Revert to Price Above Original')['percent'].to_list()

        g_inc = gb.get_group('Gross Increase')['percent'].to_list()
        n_inc = gb.get_group('Net Increase')['percent'].to_list()
        n2g = gb.get_group('Net to Gross ratio')['percent'].to_list()
        m2m = gb.get_group('Migrated to Mather')['percent'].to_list()


        ##Visualizations
        #logo = plt.imread('Mather Logo-01.png')

        #Reformat dates
        dates = [dt[-4:] for dt in dates]

        #Global formatting of plots
        mpl['lines.linewidth'] = 4
        mpl['font.family'] = 'Times New Roman'
        mpl['font.size'] = 14

        #1 figure with 2 plots
        fig, (ax1, ax2) = plt.subplots(nrows = 2, ncols = 1, sharex=False, figsize=(20,10))
        fig.suptitle(market.title(), color='#113980', size=30, x=.512)

        #Reverts Plot
        ax1.plot(dates, revert, label='Reverts', color='#7030A0')
        ax1.plot(dates, r2b, label='Revert to Below', color='#2471B3')
        ax1.plot(dates, r2o, label='Revert to Original', color='#4BB19D')
        ax1.plot(dates, r2a, label='Revert to Above', color='#113980')
        #Formatting
        ax1.legend(loc='upper right', bbox_to_anchor=(1.15, 1), ncol=1, fancybox=True, shadow=True)
        ax1.grid(True, color = '#E0E0E0')
        ax1.set_title('Reverts', fontsize=20)
        ax1.set_ylabel('Percent (%)')

        #Increases Plot
        ax2.plot(dates, g_inc, label='Gross Increase', color='#7030A0')
        ax2.plot(dates, n_inc, label='Net Increase', color='#2471B3')
        ax2.plot(dates, n2g, label='Net to Gross Ratio', color='#4BB19D')
        ax2.plot(dates, m2m, label='Migrated to Mather', color='#113980')
        #Formatting
        ax2.legend(loc='upper right', bbox_to_anchor=(1.157, 1), ncol=1, fancybox=True, shadow=True)
        ax2.grid(True, color = '#E0E0E0')
        ax2.set_title('Increases', fontsize=20)
        ax2.set_ylabel('Percent (%)')
        ax2.set_xlabel('Date (MMDD)')


        if args.output:
            plt.savefig(args.output + ('/%s.png' % market), bbox_inches='tight')
        else:
            #plt.savefig('D:/data/GateHouse/vis/%s.png' % market, bbox_inches='tight')
            plt.savefig('C:/Users/mtrinh/Desktop/gannett report pics/%s.png' % (market), bbox_inches='tight')

        plt.close('all')
