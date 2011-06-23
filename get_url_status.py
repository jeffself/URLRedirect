import sys
import httplib
import urllib2
from urlparse import urlparse
import csv
import xlrd

def get_redirected_url(url):
    opener = urllib2.build_opener(urllib2.HTTPRedirectHandler)
    request = opener.open(url)
    return request.url

if __name__ == '__main__':

    excel_file = sys.argv[1]
    
    """ Try and open the Excel file"""
    try:
        wb = xlrd.open_workbook(excel_file)
    except IOError:
        print 'cannot open', excel_file
    else:
        sh = wb.sheet_by_index(0)    
        for rownum in range(sh.nrows):
            """ Ignore first row since it has a heading """
            if rownum > 0:
                combined_url = sh.row_values(rownum)[0]
                o = urlparse(combined_url)
                conn = httplib.HTTPConnection(o.netloc)
                conn.request("GET", o.path)
                response = conn.getresponse()

                """ If http status code is a redirect (300 - 399) 
                    then we want to write it to the csv file. """
                if 300 <= response.status < 400:
                    ofile = open('urlredirect.csv', "wb")
                    writer = csv.writer(ofile, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
                    redirected_url = get_redirected_url(combined_url)
                    #redirected_url
                    row = response.status, response.reason, combined_url, redirected_url
                    writer.writerow(row)
                    ofile.close()
                conn.close()
