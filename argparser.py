import argparse
import os
from distutils.util import strtobool

################################################################
## This file is responsible for command line custom commands ##
################################################################

parser = argparse.ArgumentParser(prog="\n \n \n \n \
    ======== Students Shuffling Tool ========", 
    description='Students Shuffling Tool: \
    A tool made to read an excel file and generate an excel file with multiple sheets based on the user input \
    where the user can choose the excel file to read from, number of sections, number of stacks that he wants the shuffling for.',
    usage='%(prog)s \n \nThe tool reads and write excel files with extension of "xlsx" \
    and does not support the old "xls" version.')

def extant_file(x):
    """
    Takes x as file path checks that file exists but does not open,
    returns does not exist if not valid or not found or returns x as the path.
    """
    if not os.path.exists(x):
        raise argparse.ArgumentTypeError("{0} does not exist".format(x))
    return x

parser.add_argument('-f', '--file', default="sample1", type=str,
                    help='Enter an existing file name, \
                    by default set to "sample1"')
parser.add_argument('-p', '--path', type=extant_file, 
                    help='Enter a path file system Ex."c:\\users\\m2\\osaid\\sample.xlsx"')
parser.add_argument('-sct', '--sections', default=3, type=int,
                    help='Enter the numbers of sections to shuffle, by default set to "3"')
parser.add_argument('-st', '--stacks', nargs='+', type=str, default=['Web Fundamentals', 'Python', 'Java', 'MERN'],
                    help='Enter a list of stacks to shuffle Ex."Web Fundamentals" "Python" "Java" ...\n, \
                    by default set to ["Web Fundamentals", "Python", "Java", "MERN"]')
parser.add_argument('-gexn', '--generatedexcelname', default='output', type=str,
                    help='Enter the name of the generated excel file Ex. "test1",\n by default set to "output"')
parser.add_argument('-wclsl', '--wcolsl', nargs='+', default=[],
                    help='Enter generated column names, by default [] \
                    works only with shuffling type "1".')
parser.add_argument('-sft', '--shuffletype', default=1, type=int,
                    help='Enter a number that represents a supported shuffling option as for now \
                    there is two shuffling types "1" and "2",\n by default set to "1"')
parser.add_argument('-wst', '--withstyle', default=0, type=lambda x: bool(strtobool(x)),
                    help='Enter bool for basic styling')
parser.add_argument('-rcln', '--rcolname', default="Name", type=str,
                    help='Enter column name to read from, by default "Name"')
parser.add_argument('-rclsl', '--rcolslist', nargs='+', default=[],
                    help='Enter a list of columns to read from 2 columns max for now :((((, by default empty []')