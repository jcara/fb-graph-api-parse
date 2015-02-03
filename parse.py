#!/usr/bin/env python
import codecs
import json
import glob
import xlwt

filename = 'parsed_api_data.xls'
workbook = xlwt.Workbook()
sheet = workbook.add_sheet('Data')
author = "Example company"
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
data = "fb-graph-api-data.json"
files = glob.glob('./data/*.json')

msg = ""    
msgcount = 0
worksheet_row = 1
author_post_count = 0
for data in files:
    with open(data, "r") as f:
        result = json.loads(f.read())
    msgcount += len(result['data'])
    for i in result['data']:
        if "message" in i:
            x=1
            content = i["message"].encode('ascii','ignore')
            if i["from"]["name"]==author:
                author_post_count += 1
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
            if "comments" in i:
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

print "Saved " + str(worksheet_row-1) + " rows of data in "+ filename
style = xlwt.XFStyle()
style.alignment.wrap = 1
workbook.save(filename)
