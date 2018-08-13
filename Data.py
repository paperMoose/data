import csv
import sys
import re
import os
#-*- encoding: utf-8 -*-


class Data:

    def __init__(self):

        # fields
        self.headers = []
        self.data = []
        self.header2data = {}

    # reads in xlsx based data
    def readXlsxData(self, filename, index):
        workbook = xlrd.open_workbook(
            os.path.join("input", filename), on_demand=True)
        self.headers = None
        try:
            sheet = workbook.sheet_by_index(index)
            self.headers = sheet.row_values(0)
        except:
            if self.headers is None:
                print "Index out of bounds"
                raise StopIteration
        # decode all strings to unicode for uniformity
        self.decodeListToUnicode(self.headers)
        for rowx in range(1, sheet.nrows):
            row = sheet.row_values(rowx)
            self.data.append(row)
        for i in range(len(self.headers)):
            self.header2data[self.headers[i]] = i
        # decode all strings to unicode for uniformity
        for row in self.data:
            self.decodeListToUnicode(row)
        workbook.release_resources()
        del workbook
        print "Sheet ", index, "succesfully ingested"

    # reads the csv data from a file
    def readCsvData(self, filename):
        # read the file lines
        fp = file(os.path.join("input", filename), "rU")
        lines = fp.readlines()
        fp.close()
        # create a csv object
        csvr = csv.reader(lines)
        # set raw_headers to first line and clean
        self.headers = csvr.next()
        for cell in self.headers:
            cell = cell.strip()
        # decode all strings to unicode for uniformity
        self.decodeListToUnicode(self.headers)
        # loop through the rest of csvr and append each list to raw_data
        for thing in csvr:
            self.data.append(thing)
        # loop through the headers and k,v pair them w/ the corresponding index
        i = 0
        for cell in self.headers:
            self.header2data[self.headers[i]] = i
            i += 1
        # decode all incoming data to unicode for uniformity
        for row in self.data:
            self.decodeListToUnicode(row)

    # utility method for decoding lists of strings to unicode
    def decodeListToUnicode(self, row):
        for x in range(len(row)):
            if isinstance(row[x], str):
                row[x] = row[x].decode('utf-8', "ignore")
            if isinstance(row[x], unicode):
                pass
            else:
                row[x] = str(row[x]).decode('utf-8', "ignore")

    # returns a list of the raw headers
    def getHeaders(self):
        return self.headers

    # returns the number of raw columns
    def getNumColumns(self):
        return len(self.headers)

    # returns the number of rows
    def getNumRows(self):
        return len(self.data)

    # returns a row of raw data with the specified row number
    def getRow(self, rowNum):
        return self.data[rowNum]

    # returns a column of data with the specified header string
    def getColumn(self, header):
        # list to column values
        col = []
        # header index
        ind = self.header2data.get(header)
        # adding data to column list
        for row in self.data:
            col.append(row[ind])
        return col

    # returns the raw data at the given header, with the given row number
    def getValue(self, rowNum, header):
        return self.data[rowNum][self.header2data.get(header)]

    # sets the value at the given header, with the given row number
    def setValue(self, rowNum, header, value):
        self.data[rowNum][self.header2data.get(header)] = value

    # adds a column to the data set require a header, a type, and the correct
    # number of points
    def addColumn(self, header, plist=None):
        # adding header to list of headers
        self.headers.append(header)
        # initializing counter
        c = 0
        # loop through raw data
        for row in self.data:
            if plist is not None:
                # appending data to end of row
                row.append(plist[c])
                c += 1  # incrementing counter
            else:
                row.append("")
        # adding entry to headers2raw dictionary
        self.header2data[header] = len(self.headers) - 1

    def removeColumn(self, header):
        if header is None or header is "" or isinstance(header, str):
            print "Invalid input"
            return
        idx = self.header2data[header]
        for x in range(len(self.data)):
            row = self.data[x]
            row.pop(idx)
        self.headers.pop(idx)

    # Mapping function on the data field
    def mapData(self, function):
        for x in range(self.getNumColumns()):
            self.headers[x] = function(
                repr(self.headers[x].encode('utf-8')))
        for x in range(self.getNumRows()):
            for y in range(self.getNumColumns()):
                self.data[x][y] = function(
                    repr(self.data[x][y].encode('utf-8')))

    # saves to file
    def save(self, filename=None):
        if filename is None:
            print "No filename Provided"
            raise Exception()
        subdirectory = "results"
        try:
            os.mkdir(subdirectory)
        except Exception:
            pass
        with open(os.path.join(subdirectory, filename), 'wb') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(self.headers)
            for row in self.data:
                writer.writerow(row)

    # writes individual cells to file
    def saveCell(self, headers):
        print len(self.data)
        print self.header2data
        filenameIdx = self.header2data.get(headers[0])
        fieldIdx = self.header2data.get(headers[1])
        for x in range(len(self.data)):
            print self.data[x][filenameIdx]
            if fieldIdx is None or filenameIdx is None:
                return
            filename = "txt/" + self.data[x][filenameIdx]
            with open(filename, 'wb') as csvfile:
                csvfile.write(self.data[x][fieldIdx])

    # prints the raw data
    def printData(self):
        print self.headers
        for thing in self.data:
            print thing
