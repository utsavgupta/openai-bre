import re
import os
import sys

function_start_re = re.compile("^(Public|Private)\s(Sub|Function)\s(.+?)\(")
end_sub = "End Sub"
end_function = "End Sub"
output_dir = r"..\chunks"

def main():
    
    create_directory_if_not_exist(output_dir)
    input_dir = sys.argv[1]
    process_dir(input_dir)


def process_dir(dir) -> dict:

    files = os.listdir(dir)

    vb_files = [ f for f in files if f.endswith(".frm") or f.endswith(".bas") ]

    for file_name in vb_files:
        line_map = process_file(dir, file_name)

        for func_name in line_map.keys():

            curr_out_dir = os.path.join(output_dir, file_name)
            create_directory_if_not_exist(curr_out_dir)

            with open(os.path.join(curr_out_dir, func_name + ".txt"), "w") as f:
                f.writelines(line_map[func_name])



def process_file(dir, file_name) -> dict:
    
    line_map = dict()

    with open(os.path.join(dir, file_name)) as f_in:

        lines = f_in.readlines()
        file_line_length = len(lines)
        end_stmt = None
        curr_proc = None
        curr_line = 0

        while curr_line < file_line_length:

            if curr_proc is not None:

                line_map[curr_proc].append(lines[curr_line])
                
                if lines[curr_line].strip() == end_stmt:
                    curr_proc = None
                    end_stmt = None

                curr_line += 1
                continue

            match = function_start_re.match(lines[curr_line])

            if match is not None:
                
                proc_type = match.group(2)
                proc_name = match.group(3)

                curr_proc = proc_name
                end_stmt = get_end_statement(proc_type)

                line_map[curr_proc] = [lines[curr_line]]

            curr_line += 1

    return line_map

def get_end_statement(proc_type) -> str:

    if proc_type == "Function":
        return "End Function"
    else:
        return "End Sub"

def create_directory_if_not_exist(dir):

    if not os.path.isdir(dir):
        os.mkdir(dir)

if __name__ == '__main__':
    main()