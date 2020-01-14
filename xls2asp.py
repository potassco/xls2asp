"""
Converts an instance given as a set of excel tables
into a set of asp facts.

Input: Excel xlsx file

Output: Logic program instance file
"""

import warnings
import csv
import argparse
import sys
import traceback
import openpyxl as xls
import math
import warnings
import re
import datetime
from operator import itemgetter

current_datetime = datetime.datetime.now()
current_date = datetime.date(current_datetime.year, current_datetime.month, current_datetime.day)

# list all styles and types 
list_of_styles = ["matrix_xy", "row"]
list_of_types = ["auto_detect", "skip", "int", "constant", "time", "date", "date_rel", "datetime", "datetime_rel", "string"]

def write_category_comment(output, pred):
    output.write('%' * (len(pred) + 6) + '\n')
    output.write('%% ' + pred + ' %%\n')
    output.write('%' * (len(pred) + 6) + '\n')

class Xls2AspError(ValueError):
    def __init__(self, msg, sheet = "?", cell = (1,0)):
        super(Xls2AspError, self).__init__(msg)
        self.sheet = sheet
        self.cell  = cell

class SheetRowColumnWrongTypeValueError(ValueError):
    def __init__(self, table, row, col, msg, value=None):
      ValueError.__init__(self, 'Wrong type in sheet "{}" row "{}" column "{}": {}'.format(table, row, col, msg), value)

class Conversion:

    @staticmethod
    def datetime2min(value):
        delta = current_datetime - value
        return str(delta.days)+"*24*60+"+str(int(value.second/60))

    @staticmethod
    def date2min(value):
        delta = current_date - value
        return str(delta.days)+"*24*60"
    
    @staticmethod
    def date2tuple(value):
        return "("+str(value.day)+","+str(value.month)+","+str(value.year)+")"

    @staticmethod
    def datetime2tuple(value):
        return "("+str(value.day)+","+str(value.month)+","+str(value.year)+","+str(value.hour)+","+str(value.minute)+","+str(value.second)+")"

    @staticmethod
    def is_int(value):
        return Conversion.is_single_int(value) or Conversion.is_set_of_int(value)

    @staticmethod
    def is_single_int(value):
        try:
            return int(value) == float(value)
        except (TypeError, ValueError, AttributeError):
            return False 

    @staticmethod
    def is_set_of_int(value):
        try:
            for i in value.split(";"):
                a = int(i) == float(i)
            return True 
        except (TypeError, ValueError, AttributeError):
            return False
        return False

    @staticmethod
    def normalize_int(value):
        if Conversion.is_single_int(value):
            return value
        else:
            return "("+value+")"

    @staticmethod
    def is_unsigned(value):
        try:
            return int(value) == float(value) &int(value) >= 0
        except (TypeError, ValueError):
            return False

    @staticmethod
    def make_predicate(value):
        value = ''.join(e for e in value if e.isalnum())
        if Conversion.is_int(value[0]):
            value = "pred"+value
        if value[0].isupper(): 
            value = value[0].lower()+value[1:]
        return value.replace(" ","_")

    @staticmethod
    def make_asp_constant(value):
        """
        ensures lowercase, no leading or trailing blanks and no whitespace in between
        """
        return int(value) if Conversion.is_int(value) else str(value).strip().lower().replace(" ","_")

    @staticmethod
    def is_constant_char(value):
        try:
            if value.isalnum() or value == "_":
                return True
            else:
                return False
        except (TypeError, ValueError, AttributeError):
            return False


    @staticmethod
    def is_single_constant(value):
        """
        ensures lowercase and no leading or trailing blanks, no whitespace in between
        """
        try:
            for i in value:
                if not Conversion.is_constant_char(i):
                    return False
            for i in value:
                if i.islower():
                    return True
                elif i != "_":
                    return False
        except (TypeError, ValueError):
            return False

    @staticmethod
    def is_set_of_constant(value):
        """
        ensures lowercase and no leading or trailing blanks, no whitespace in between
        """
        try: 
            for i in value.split(";"):
                if not Conversion.is_single_constant(i):
                    return False
            return True
        except (TypeError, ValueError, AttributeError):
            return False

    @staticmethod
    def is_asp_constant(value):
        """
        ensures lowercase and no leading or trailing blanks, no whitespace in between
        """
        return Conversion.is_single_constant(value) or Conversion.is_set_of_constant(value)

    def normalize_constant(value):
        if Conversion.is_single_constant(value):
            return value
        else:
            return "("+value+")"

    _regtime   = re.compile('(\d\d):(\d\d):00')
    @staticmethod
    def is_hhmm00(value):
        """
        ensures a string of the format HH:MM:00
        """
        try:
            return Conversion._regtime.match(value) != None
        except (TypeError, ValueError):
            return False
    
    @staticmethod
    def hhmm00_in_m(value):
        """
        convert a string of the format HH:MM:00 to minutes
        """
        x = Conversion._regtime.match(value)
        return int(x.group(1))*60+int(x.group(2))

    @staticmethod
    def make_asp_tuple(value):
        "converts comma-separated string into a tuple of asp constants"
        if value == None: return tuple()
        return tuple([Conversion.make_asp_constant(x) for x in str(value).split(',')])

    @staticmethod
    def make_int(value, min = None, max = None, msg = None):
        x = int(value)
        if (min != None and x < min) or (max != None and x > max): raise ValueError("int out of range" if not msg else msg)
        return x
    @staticmethod
    def make_unsigned(value):
        return Conversion.make_int(value, min=0)
    @staticmethod
    def make_percent(value):
        return Conversion.make_int(value, min=0, max=100)
    @staticmethod
    def make_minutes_from_hours(hours, max = None):
        minutes = float(hours) * 60
        if not Conversion.is_int(minutes): raise ValueError("Time conversion loses precision")
        if minutes < 0 or (max != None and minutes > float(max)*60): raise ValueError("Time out of range")
        return int(minutes)

    @staticmethod
    def make_minutes_from_datetime(time):
        return Conversion.make_minutes_from_hours(time.hour) + time.minute

class Template:
    """
    Class for maintaining template
    """
    def __init__(self):
        self.template = {}

    def read(self, fileName):
        with open(fileName,"r") as f:
            for line in f:
                reader = csv.reader([line], skipinitialspace=True)
                template = {}
                line = next(reader)
                table = line[0]
                self.add_table(table)
                style = line[1].strip()
                if style not in list_of_styles:
                    raise ValueError('style not valid: '+style)
                self.add_style(table, style)
                types=line[2:]
                if style == "matrix_xy":
                    if len(types) != 3:
                        raise ValueError('3 types are needed to read in matrix style')
                for t in types:
                    types[types.index(t)] = t.strip()
                    t = t.strip()
                    if t not in list_of_types:
                        raise ValueError('type not valid: '+t)
                self.add_types(table, types)
            f.close()

    def add_table(self, table):
        """
        Adds a table and ensures it is unique
        """
        assert table not in self.template, "Duplicate table '%r' in template" % table
        self.template.setdefault(table,{})

    def add_types(self, table, types):
        """
        Adds a predicate types to a table
        """
        self.template.setdefault(table,{}).setdefault("types", types)

    def add_style(self, table, style):
        self.template.setdefault(table,{}).setdefault("style", style)

class Instance:
    """
    Class for maintaining data of an instance file
    """
    def __init__(self, template):
        self.data = {}
        self.template = template

    def write_table_row_style(self,table,file):
        """
        Writes table content to facts row by row
        """
        write_category_comment(file,table)
        predName = Conversion.make_predicate(table)
        for row in self.data[table]["rows"]:
            pred = predName+'('
            for value in self.data[table]["rows"][row]:
                pred += str(value)+','
                #pred += str(self.data[table]["rows"][i])+','
            pred = pred[0:len(pred)-1]
            pred += ').\n'
            file.write(pred)
        file.write('\n')
        file.write('\n')

    def write_table_matrix_xy_style(self,table,file):
        """
        Writes table content to facts
        """
        write_category_comment(file,table)
        predName = Conversion.make_predicate(table)
        for r in range(2, len(self.data[table]["rows"])+1): 
            y = self.data[table]["rows"][r][0]
            for i in range(1, len(self.data[table]["rows"][r])): 
                if self.data[table]["rows"][r][i] != None:
                    pred = predName+'('+str(self.data[table]["rows"][1][i])+','
                    pred += str(y)+','
                    pred += str(self.data[table]["rows"][r][i])+').\n'
                    file.write(pred)
        file.write('\n')
        file.write('\n')

    def write_predicate(self,predicate,instances,file):
        """
        Writes predicate instances
        """
        for value in instances:
            file.write(Conversion.make_predicate(predicate)+'('+str(value[0])+','+str(value[1])+').\n')
        file.write('\n')

    def write(self, file):
        for table in self.data:
            style = self.data[table]["style"]
            if style == "row":
                self.write_table_row_style(table,file)
            if style == "matrix_xy":
                self.write_table_matrix_xy_style(table,file)

    def add_table(self, table):
        """
        Adds a table and ensures it is unique
        """
        assert table not in self.data, "Duplicate table '%r'" % table
        self.data.setdefault(table,{})

    def add_style(self, table, style):
        """
        Adds style to a table
        """
        self.data.setdefault(table,{}).setdefault("style",style)

    def add_row(self, table, id, row):
        self.data.setdefault(table,{}).setdefault("rows",{}).setdefault(id,row)

    def get_test(self, type):
        if type == "int":
            return self.test_int
        elif type == "constant":
            return self.test_constant
        elif type == "string":
            return self.test_string
        elif type == "time":
            return self.test_time
        elif type == "date":
            return self.test_date
        elif type == "date_rel":
            return self.test_date_rel
        elif type == "datetime":
            return self.test_datetime
        elif type == "datetime_rel":
            return self.test_datetime_rel
        elif type == "skip":
            return None
        elif type == "auto_detect":
            return self.test_auto_detect
        else:
            raise ValueError('Type not valid: '+type)

    def test_string(self, table, row, col, value):
        if isinstance(value,str): 
            return "\""+value+"\""
        else:
            if isinstance(value,unicode) :
                return "\""+value+"\""
            print(value+" is not a string")
            print(isinstance(value,str)) 
            raise SheetRowColumnWrongTypeValueError(table, row, col, "Expecting a string, getting:", value)

    def test_int(self, table, row, col, value):
        if Conversion.is_int(value):
            return Conversion.normalize_int(value)
        else:
            raise SheetRowColumnWrongTypeValueError(table, row, col, "Expecting an int, getting:", value)

    def test_constant(self, table, row, col, value):
        if Conversion.is_asp_constant(value):
            return Conversion.normalize_constant(value)
        else:
            raise SheetRowColumnWrongTypeValueError(table, row, col, "Expecting a constant, getting:", value)

    def test_time(self, table, row, col, value):
        if isinstance(value, datetime.time):
            return str(value.hour)+"*60+"+str(value.minute)
        else:
            raise SheetRowColumnWrongTypeValueError(table, row, col, "Expecting a time, getting:", value)

    def test_datetime(self, table, row, col, value):
        if isinstance(value, datetime.datetime):
            return Conversion.datetime2tuple(value) 
        else:
            raise SheetRowColumnWrongTypeValueError(table, row, col, "Expecting a datetime, getting:", value)

    def test_datetime_rel(self, table, row, col, value):
        if isinstance(value, datetime.datetime):
            return Conversion.datetime2(value) 
        else:
            raise SheetRowColumnWrongTypeValueError(table, row, col, "Expecting a datetime, getting:", value)

    def test_date(self, table, row, col, value):
        if isinstance(value, datetime.date): 
            return Conversion.date2tuple(value)
        else:
            raise SheetRowColumnWrongTypeValueError(table, row, col, "Expecting a date, getting:", value)

    def test_date_rel(self, table, row, col, value):
        if isinstance(value, datetime.date): 
            return Conversion.date2min(datetime.date(value.year, value.month, value.day))
        else:
            raise SheetRowColumnWrongTypeValueError(table, row, col, "Expecting a date, getting:", value)

    def test_auto_detect(self, table, row, col, value):
        if Conversion.is_int(value):
            return Conversion.normalize_int(value)
        elif Conversion.is_asp_constant(value):
            return Conversion.normalize_constant(value)
        elif isinstance(value,datetime.time):
            return str(value.hour)+"*60+"+str(value.minute)
        elif isinstance(value, datetime.datetime):
            return Conversion.datetime2tuple(value) 
        elif isinstance(value, datetime.date):
            return Conversion.date2tuple(value) 
        elif isinstance(value,str): 
            return "\""+value+"\""
        elif value == None:
            return None
        else: 
            raise SheetRowColumnWrongTypeValueError(table, row, col, "Unknown type for ", value)

    def correct(self):
        # remove leading or trailing blanks from each value in every table
        for table in self.data:
            for r in self.data[table]["rows"]:
                row = self.data[table]["rows"][r]
                for value in row:  
                    try:
                        row[row.index(value)] = value.strip()
                    except (AttributeError):
                        pass
        for table in self.data:
            style = self.template[table]["style"]
            if style == "row":
                self.correct_row_style(table)
            elif style == "matrix_xy":
                self.correct_matrix_xy_style(table)
            else:
                raise ValueError('style not valid: '+style)

    def correct_row_style(self, table):
        unexpected = 0
        nb_col = len(self.template[table]["types"])
        self.data[table]["rows"].pop(1)#ignore first line  
        #ignoring empty rows
        list_empty = []
        for row in self.data[table]["rows"]:
            empty = 1
            for value in self.data[table]["rows"][row]:
                if value != None:
                    empty = 0
            if empty == 1:
                list_empty.append(row)
        for row in list_empty:
            self.data[table]["rows"].pop(row)
        for row in self.data[table]["rows"]:
            if len(self.data[table]["rows"][row]) > nb_col:
                unexpected = 1
                self.data[table]["rows"][row] = self.data[table]["rows"][row][0:nb_col]
        if unexpected:
            sys.stderr.write("WARNING: Undefined column in sheet \""+table+"\", ignoring it\n")
        col = 0
        for type in self.template[table]["types"]:
            if type == "skip":
                for row in self.data[table]["rows"]:
                    self.data[table]["rows"][row].pop(col)
            else:
                test = self.get_test(type)
                for row in self.data[table]["rows"]:
                    value = self.data[table]["rows"][row][col]
                    if value == None:
                        raise SheetRowColumnWrongTypeValueError(table, row, col, '"None" is not an expected value.')
                    self.data[table]["rows"][row][col] = test(table, row, col+1, value)
                col += 1

    def correct_matrix_xy_style(self, table):
        type_x = self.template[table]["types"][0]
        type_y = self.template[table]["types"][1]
        type_v = self.template[table]["types"][2]
        
        self.remove_empty_column(table)

        # test type for x (= first line)
        test = self.get_test(type_x)
        row_x = self.data[table]["rows"][1]
        #for i in range(1,len(row_x)): 
        for i in range(1,len(row_x)): 
            row_x[i] = test(table, 1, i+1, row_x[i])

        # test type for y (= first column)
        test = self.get_test(type_y)
        for r in range(2, len(self.data[table]["rows"])+1): 
            self.data[table]["rows"][r][0] = test(table, r, 1, self.data[table]["rows"][r][0])
        
        # test type for the inner matrix
        test = self.get_test(type_v)
        for r in range(2, len(self.data[table]["rows"])+1): 
            for i in range(1, len(self.data[table]["rows"][r])): 
                self.data[table]["rows"][r][i] = test(table, r, i+1, self.data[table]["rows"][r][i])

    def get_table_style(self, table):
        if table not in self.template:
            sys.stderr.write("WARNING: Sheet \""+table+"\" is not defined in the template\n")
            return "skip"
        style = self.template[table]["style"]
        return style

    def remove_empty_column(self, table):
        empty_columns = [] 
        for col in range(len(self.data[table]["rows"][1])):
            empty = True 
            for row in self.data[table]["rows"]:
                if self.data[table]["rows"][row][col] != None:
                    empty = False 
                    break
            if empty:
                empty_columns.append(col)
        for i in range(len(empty_columns)):
            sys.stderr.write("WARNING: Column "+str(empty_columns[i])+" in sheet \""+table+"\" is empty, ignoring it\n")
            for row in self.data[table]["rows"]:
                self.data[table]["rows"][row].pop(empty_columns[i]-i)




def compile_regex_from_list(values):
    reg = '({})$'.format('|'.join(map(lambda x:'({})'.format(x), values)))
    return re.compile(reg)

class XlsReader:
    def __init__(self, instance):
        # Expected worksheets xlsx file and their parsing functions
        self.instance = instance
        self.active_cell = (1,0)

    def get_cell_value(self, cell, inc = (0,1)):
        if   getattr(cell, "col_idx", None): self.active_cell = (cell.row, cell.col_idx)
        elif getattr(cell, "column", None):  self.active_cell = (cell.row, cell.column)
        else: self.active_cell = (self.active_cell[0] + inc[0], self.active_cell[1] + inc[1])
        return cell.value

    def parse_row(self, row, first = 0):
        cols = []
        for i in range(first, len(row)):
            cols.append(row[i].value)
        return cols

    def parse_data_row(self, head, row, first=0, skip=True):
        assert len(row) >= len(head), "Unexpected number of columns %r - expected %r" % (len(row), len(head))
        for (x,y) in zip(head, row[first:]):
            if y.value == None and skip: continue
            yield (x, self.get_cell_value(y))

    def parse_data(self, data, seperator=";"):
        for val in str(data).split(seperator):
            if Conversion.is_hhmm00(val):
                yield Conversion.hhmm00_in_m(val)
            elif Conversion.is_int(val):
                yield int(val)
            elif Conversion.is_asp_constant(val):
                yield val
            else:
                yield '"' + val.replace("\"","\\\"") + '"'

    def parse_table(self, sheet, style):
        table = sheet.title
        print("Parsing sheet \""+table+"\" with style \""+style+"\"")
        self.instance.add_table(sheet.title)
        self.instance.add_style(sheet.title,style)
        self.active_cell  = (1,0)
        self.active_sheet = sheet
        try:
            id = 1
            for r in sheet.iter_rows(min_row=1):
                row = self.parse_row(r)
                self.instance.add_row(table, id, row)
                id += 1
        except Exception as e:
            raise Xls2AspError(str(e), self.active_sheet, self.active_cell)


    def parse(self,input):
        """
        Parses input excel table
        """
        wb = xls.load_workbook(input, read_only=True, data_only=True)
        if self.__update_dimensions(wb):
            wb.close()
            wb = xls.load_workbook(input, read_only=False, data_only=True)
        for sheet in wb:
            style = self.instance.get_table_style(sheet.title)
            if style == "skip":
                print("Skipping sheet: "+sheet.title)
            else:
                self.parse_table(sheet,style)
        for table in self.instance.template:
            if table not in self.instance.data:
                raise ValueError("Sheet \""+table+"\" not found") 
            if not self.instance.data[table].__contains__("rows"):
                sys.stderr.write("WARNING: Sheet \""+table+"\" is empty, ignoring it\n")
                self.instance.data.pop(table)
                print("Skipping sheet: "+table)

    def __update_dimensions(self, workbook):
        for sheet in workbook.worksheets:
            if sheet.max_column > 50 or sheet.max_row > 1000:
                return True
        return False

def main():
    try:
        parser = argparse.ArgumentParser(
                description="Converts an input table to facts"
                )
        parser.add_argument('--output', '-o', metavar='<file>',
                help='Write output into %(metavar)s', default = sys.stdout, required=False)
        parser.add_argument('--xls', '-x', metavar='<file>',
                help='Read xls file from %(metavar)s', required=True)
        parser.add_argument('--template', '-t', metavar='<file>',
                help='Read template from %(metavar)s', required=True)

        args = parser.parse_args()
        template = Template()
        template.read(args.template)
        instance = Instance(template.template)
        reader = XlsReader(instance)
        reader.parse(args.xls)
        instance.correct()
        if args.output ==  sys.stdout:
            instance.write(args.output)
        else:
            with open(args.output, 'w') as f:
                instance.write(f)
        return 0
    except Xls2AspError as e:
        sys.stderr.write("*** Exception: {}\n".format(e))
        sys.stderr.write("***   In sheet={0}:{1}{2}\n".format(e.sheet, utils.cell.get_column_letter(e.cell[1]), e.cell[0]))
        return 1
    except Exception as e:
        traceback.print_exception(*sys.exc_info())
        return 1

if __name__ == '__main__':
    sys.exit(main())

