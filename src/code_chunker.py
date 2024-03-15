import re
import os
import sys

function_start_re = re.compile("^(Public|Private)\s(Sub|Function)\s.(.+?)\(")
output_dir = r"..\chunks"

def main():
    
    create_directory_if_not_exist(output_dir)
    input_dir = sys.argv[1]


def process_dir(dir):

    files = os.listdir(dir)

    vb_files = [ f for f in files if f.endswith(".frm") or f.endswith(".bas") ]

    for file in vb_files:
        process_file(dir, file)

def process_file(dir, file_name):
    pass

def create_directory_if_not_exist(dir):

    if not os.path.isdir(dir):
        os.mkdir(dir)

if __name__ == '__main__':
    main()