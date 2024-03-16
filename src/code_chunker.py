import re
import os
import sys
import yaml

function_start_re = re.compile("^(Public|Private)\s(Sub|Function)\s(.+?)\(")
control_re = re.compile("^Begin\s(.+)\s(.+)$")
end_sub = "End Sub"
end_function = "End Sub"
output_dir = r"..\chunks"

class ProcessOutput:

    def __init__(self, controls, locs):
        self._controls = controls
        self._locs = locs
    
    def controls(self) -> dict:
        return self._controls

    def locs(self) -> dict:
        return self._locs

def process_dir(dir: str) -> dict:

    files = os.listdir(dir)

    vb_files = [ f for f in files if f.endswith(".frm") or f.endswith(".bas") ]

    for file_name in vb_files:
        
        process_output = process_file(dir, file_name)
        line_map = process_output.locs()
        controls = process_output.controls()

        curr_out_dir = os.path.join(output_dir, file_name)
        create_directory_if_not_exist(curr_out_dir)

        if controls is not None:
            with open(os.path.join(curr_out_dir, "controls.yml"), "w") as f:
                f.write(yaml.dump(controls))

        for func_name in line_map.keys():

            with open(os.path.join(curr_out_dir, func_name + ".txt"), "w") as f:
                f.writelines(line_map[func_name])



def process_file(dir: str, file_name: str) -> ProcessOutput:

    with open(os.path.join(dir, file_name)) as f_in:

        lines = f_in.readlines()
        controls = None

        if file_name.endswith(".frm"):
            controls = process_controls(lines)
        
        locs =  process_locs(lines)

        return ProcessOutput(controls, locs)
       

def process_controls(lines: list[str]) -> dict:
    
    controls_map = dict()
    control_stk = list()
    have_seen_form = False

    for line in lines:
        
        if have_seen_form and len(control_stk) == 0:
            break
        
        match = control_re.match(line.strip())

        if match is not None:

            control_name = match.group(2)
            control_type = match.group(1)

            if have_seen_form is False:
                controls_map[control_name] = { "type": control_type, "controls": list() }
                control_stk.append(controls_map[control_name]["controls"])
                have_seen_form = True 
                continue
            
            new_control = { control_name: { "type": control_type, "controls": list() } }
            control_stk[-1].append(new_control)
            control_stk.append(new_control[control_name]["controls"])

        if line.strip() == "End":
            control_stk = control_stk[:-1]

    return controls_map

def process_locs(lines: list[str]) -> dict:
    
    line_map = dict()
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

def get_end_statement(proc_type: str) -> str:

    if proc_type == "Function":
        return "End Function"
    else:
        return "End Sub"

def create_directory_if_not_exist(dir: str):

    if not os.path.isdir(dir):
        os.mkdir(dir)

def main():
    
    create_directory_if_not_exist(output_dir)
    input_dir = sys.argv[1]
    process_dir(input_dir)


if __name__ == '__main__':
    main()