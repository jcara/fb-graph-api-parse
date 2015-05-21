#!/usr/bin/env python
import os
import sys
import codecs
import datetime
import json
import glob
import string
import xlwt
import getopt
import urllib2
import urlparse
import md5

from variables import access_token

def write_statistics(filename, weekday_post_count=None, hour_post_count=None):

    statistics_file = open(filename + "_statistics.txt","w")
    statistics_file.write("Statistics of parsed Facebook Graph API data\n")
    statistics_file.write("Date: " + datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S:%f") + "\n\n")
    weekdays = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    if(weekday_post_count != None):
        i = 0
        statistics_file.write("Post counts separated to weekdays: \n")
        for i in range(0,7):
            statistics_file.write(weekdays[i] + ": " + str(weekday_post_count[i]) + "\n")
            i = i + 1

    if(hour_post_count != None):
        h = 0
        statistics_file.write("Post counts separated to hours of day:\n")
        for count in hour_post_count:
            if len(str(h))==1:
                hour = str(0) + str(h)
            else:
                hour = str(h)
            statistics_file.write(hour+": "+str(count)+"\n")
            h = h + 1
    statistics_file.close() 

def convert_input_time_to_epoch(input_time):
    try:
        if(len(input_time) <= 10):
            dt = datetime.datetime.strptime(input_time, '%Y-%m-%d')
        else:
            dt = datetime.datetime.strptime(input_time, '%Y-%m-%d %H:%M:%S')
        d = int((dt - datetime.datetime(1970,1,1)).total_seconds())
    except TypeError:
        try:
            d = int(input_time)

        except TypeError:
            # if fails for both conversions, print usage and quit
            usage()
    return d


def convert_api_date_to_datetime(date_as_string):
    date_parse = date_as_string.split("-",3)
    year = int(date_parse[0])
    month = int(date_parse[1])
    day = int(date_parse[2].split("T",1)[0])

    date_parse_hhmmss = date_parse[2].split("T",4)[1]
    date_parse_hhmmss = date_parse_hhmmss.split(":",3)
    hours = int(date_parse_hhmmss[0])
    minutes = int(date_parse_hhmmss[1])
    seconds =int( date_parse_hhmmss[2].split("+",1)[0])
    return datetime.datetime(year,month,day,hours,minutes,seconds)


def usage():
    file = open('usage.txt','r')
    print file.read()
    sys.exit()


def is_json(data): 
    with open(data,"r") as f:
        # If file is valid JSON, exception does not trigger
        try:
            result = json.loads(f.read())
        except ValueError, e:
            print data + " is not valid JSON."
            return False
        return True

def main(argv):
    filename = 'parsed_api_data'
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('Data')
    debug = False
    page = None 
    start_date = None
    end_date = None
    date_filter = False
    # Include comments
    include_comments = False

    # Include posts only from page itself
    page_only = True

    try:
        opts,args = getopt.getopt(argv, "hcpdo:s:e:", ["help", "comments", "page=", "debug", "output=", "start-date=", "end-date=", "page-only" ])
    except getopt.GetoptError as err:
        print(err)
        usage()

    for opt,arg in opts:
        if opt in ("h","--help"):
            usage()
        elif opt in ("c","--comments"):
            include_Comments = True
        elif opt in ("p","--page"):
            page = arg
        elif opt in ("--page-only"):
            page_only = True
        elif opt in ("d","--debug"):
            debug = True
        elif opt in ("o","--output"):
            if(len(str(arg)) > 0):
                filename = str(arg)
        elif opt in ("-s", "--start-date"):
            start_date = convert_input_time_to_epoch(arg)
        elif opt in ("-e", "--end-date"):
            end_date = convert_input_time_to_epoch(arg)
        else:
            if page == None:
                page = opt

    if page == None:
        print "No page specified."
        usage()

    columns = ['Post ID',
              'Type',
              'Author',
              'Time',
              'Content',
              'Like count',
              'Share count',
              'Comment count']

    for i, col in enumerate(columns):
        sheet.write(0,i,col)

    msg = ""    
    msgcount = 0
    worksheet_row = 1
    page_post_count = 0
    valid_files = []
    # Post count for weekdays (starting from Monday)
    weekday_post_count = [0,0,0,0,0,0,0]
    hour_post_count = []
    i = 0
    for i in range(0,24):
        hour_post_count.append(0)
     
    url = "https://graph.facebook.com/"+page+"/feed?access_token="+access_token+"&limit=10"
    if(start_date != None):
        url += "&since=" + str(start_date)
    if (end_date != None):
        url += "&until=" + str(end_date)

    print url
    page_id = None
    data_collect_complete = False
    index = 0
    handled_post_count = 0
    while data_collect_complete == False:
        index = index + 1
        d = urllib2.urlopen(url).read()
        result = json.loads(d)

        if page_id == None:
            page_id = result['paging']['next'].split('/')[4]
        
        
        print "new url read" + url

        if len(result['data']) == 0:
            data_collect_complete = True
            print "result['data'] is zero, lets quit"
            break
       
        for i in result['data']:
            
            print str(index)+" "+str(i["created_time"])
            handled_post_count = handled_post_count + 1
            
            if "message" in i:
                content = i["message"].encode('ascii','ignore')
                if i["from"]["id"]==page_id:
                    page_post_count += 1
                else:
                    if i["from"]["name"]=="Uppsala universitet":
                        print "upp:"+str(i["from"]["id"])
                    if page_only == True:
                        # Skip message
                        continue
                
                date_parse = convert_api_date_to_datetime(i["created_time"])
                weekday_of_date = date_parse.weekday()
                hour_of_date = int(date_parse.strftime("%H"))
                weekday_post_count[weekday_of_date] = weekday_post_count[weekday_of_date] + 1
                hour_post_count[hour_of_date] = hour_post_count[hour_of_date] + 1

                # Write statuses
                sheet.write(worksheet_row, 0, i["id"])
                sheet.write(worksheet_row, 1, 'Post')
                sheet.write(worksheet_row, 2, i["from"]["name"])
                sheet.write(worksheet_row, 3, i["created_time"])
                sheet.write(worksheet_row, 4, i["message"])
                if "likes" in i:
                    sheet.write(worksheet_row,5,len(i["likes"]["data"]))
      
                if "shares" in i:
                    sheet.write(worksheet_row,6,i["shares"]["count"])
                
                if ("comments" in i) and include_comments == True:
                    sheet.write(worksheet_row,7,len(i["comments"]["data"]))
                
                    for ii in i["comments"]["data"]:
                        # Write comments
                        worksheet_row += 1
                        sheet.write(worksheet_row, 0, i["id"])
                        sheet.write(worksheet_row, 1, 'Comment')
                        sheet.write(worksheet_row, 2, ii["from"]["name"])
                        sheet.write(worksheet_row, 3, ii["created_time"])
                        sheet.write(worksheet_row, 4, ii["message"])
                        sheet.write(worksheet_row, 5, ii["like_count"])

                worksheet_row +=1


        url = result['paging']['next']
        if start_date != None:
            url += "&since=" + str(start_date)
            print url

    style = xlwt.XFStyle()
    style.alignment.wrap = 1

    folder = "data"
    if not os.path.exists(folder):
        os.makedirs(folder)
    workbook.save(folder + '/' + filename + ".xls")
    print "workbook saved"
    write_statistics(folder + '/' + filename, weekday_post_count, hour_post_count)

    print "Saved " + str(worksheet_row-1) + " rows of data in "+ folder +'/' + filename + ".xls."
    print "Statistics available in " + folder + '/' + filename + ".txt."

if __name__ == "__main__":
    main(sys.argv[1:])
