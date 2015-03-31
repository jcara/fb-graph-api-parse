#!/usr/bin/env python
import sys
import codecs
import datetime
import json
import glob
import string
import xlwt
import getopt

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
    #seconds = int(seconds.split(":",1)[0])
    # cut hh:mm:ss part of timestamp, we don't need it for weekday check
    
    #post_datetime = datetime.datetime(int(date_parse[0]),int(date_parse[1]),int(date_parse[2]))
    return datetime.datetime(year,month,day,hours,minutes,seconds)




def usage():
    file = open('usage.txt','r')
    print file.read()

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
    author = "Not set"
    
    # Include comments
    include_comments = False

    # Include posts only from page itself
    author_only = True

    try:
        opts,args = getopt.getopt(argv, "hcoa:dp", ["help", "comments", "only-author", "author=", "debug", "prefix="])
    except getopt.GetoptError as err:
        print(err)
        usage()
        sys.exit()

    for opt,arg in opts:
        if opt in ("h","--help"):
            f = open("usage.txt","r")
            print f.read()
            sys.exit()
        elif opt in ("c","--comments"):
            include_Comments = True
        elif opt in ("a","--author"):
            author = arg
        elif opt in ("o","--only-author"):
            author_only = True
        elif opt in ("d","--debug"):
            debug = True
        elif opt in ("p","--prefix"):
            if(len(str(arg)) > 0):
                filename = str(arg)
        else:
            print ""

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

    # Search for all files in data/ folder
    files = glob.glob('./data/*.*')
    msg = ""    
    msgcount = 0
    worksheet_row = 1
    author_post_count = 0
    valid_files = []
    # Post count for weekdays (starting from Monday)
    weekday_post_count = [0,0,0,0,0,0,0]
    hour_post_count = []
    i = 0
    for i in range(0,24):
        hour_post_count.append(0)
    # Check for validity
    for data in files:
        if is_json(data):
            valid_files.append(data)

    for data in valid_files:
        with open(data, "r") as f:
            result = json.loads(f.read())
        msgcount += len(result['data'])
        for i in result['data']:

            if "message" in i:
                content = i["message"].encode('ascii','ignore')

                if i["from"]["name"]==author:
                    author_post_count += 1
                else:
                    if author_only == True:
                        continue
                
                date_parse = convert_api_date_to_datetime(i["created_time"])
                weekday_of_date = date_parse.weekday()
                hour_of_date = int(date_parse.strftime("%H"))
                weekday_post_count[weekday_of_date] = weekday_post_count[weekday_of_date] + 1
                hour_post_count[hour_of_date] = hour_post_count[hour_of_date] + 1

                # Write statuses
                sheet.write(worksheet_row,0,i["id"])
                sheet.write(worksheet_row,1,'Post')
                sheet.write(worksheet_row,2,i["from"]["name"])
                sheet.write(worksheet_row,3,i["created_time"])
                sheet.write(worksheet_row,4,i["message"])
                if "likes" in i:
                    sheet.write(worksheet_row,5,len(i["likes"]["data"]))
      
                if "shares" in i:
                    sheet.write(worksheet_row,6,i["shares"]["count"])
                
                if ("comments" in i) and include_comments == True:
                    sheet.write(worksheet_row,7,len(i["comments"]["data"]))
                
                    for ii in i["comments"]["data"]:
                        # Write comments
                        worksheet_row += 1
                        sheet.write(worksheet_row,0,i["id"])
                        sheet.write(worksheet_row,1,'Comment')
                        sheet.write(worksheet_row,2,ii["from"]["name"])
                        sheet.write(worksheet_row,3,ii["created_time"])
                        sheet.write(worksheet_row,4,ii["message"])
                        sheet.write(worksheet_row,5,ii["like_count"])

                worksheet_row += 1

    style = xlwt.XFStyle()
    style.alignment.wrap = 1

    workbook.save(filename + ".xls")
    write_statistics(filename, weekday_post_count, hour_post_count)

    print "Saved " + str(worksheet_row-1) + " rows of data in "+ filename + ".xls."
    print "Statistics available in " + filename + ".txt."

if __name__ == "__main__":
    main(sys.argv[1:])
