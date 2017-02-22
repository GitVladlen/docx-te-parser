import sys
import os
import argparse

from teparser import process

# -------------------- STUFF FOR EXECUTE FROM PACKAGE (AS MAIN) --------------------
def process_main(docx_file_name, destination_dir):
    # define filenames
    texts_filename = os.path.join(destination_dir, "TextEncounterTexts.xml")
    register_filename = os.path.join(destination_dir, "te-register.txt")
    # define string formats for writing register file as table
    register_title_format = "|{ID:^8}|{Module:^16}|{TypeName:^20}|\n"
    register_row_format = "|{ID:>8}|{Module:>16}|{TypeName:>20}|\n"
    register_line_string = "=" * (8 + 16 + 20 + 4) + "\n"
    # do job, gain result scrips and texts
    scripts, texts = process(docx_file_name)
    # write texts
    with open(texts_filename, "w") as f:
            f.write(texts)
    # write register and scrips
    with open(register_filename, "w") as register:
        # write register title
        register.write(register_line_string)
        register.write(register_title_format.format(ID="ID", Module="Module", TypeName="TypeName"))
        register.write(register_line_string)
        # write each script to corresponding file
        for r, (te_id, text) in enumerate(scripts, start=1):
            type_name = "TextEncounter{}".format(te_id)
            generateFilename = os.path.join(destination_dir, type_name + ".py")
            # write script
            with open(generateFilename, "w") as f:
                f.write(text)
            # write record about script in register
            register.write(register_row_format.format(ID=te_id, Module="TextEncounter", TypeName=type_name))
            pass
        pass
        # write line in end of register
        register.write(register_line_string)
    pass

def process_args():
    parser = argparse.ArgumentParser(description="Export Text Encounter from docx file to Mengine")
    parser.add_argument("--docx", help="path of the docx file")
    parser.add_argument("--dest", help="path of destination directory")

    args = parser.parse_args()
    return args

# ---------------------------------------------------------------------------------------------
if __name__ == '__main__':
    # setup default file names
    cur_dir_path = os.path.dirname(__file__)
    destination_dir = os.path.join(cur_dir_path, "debug/")
    docx_file_name = os.path.join(destination_dir, "te-format.docx")
    # check console args
    args = process_args()
    if args.docx and args.dest:
        docx_file_name = args.docx
        destination_dir = args.dest
        pass
    # check docx file existance
    if not os.path.exists(docx_file_name):
        print("File {} dose not exist.".format(docx_file_name))
        sys.exit(1)
        pass
    # check destination dir existance
    if not os.path.exists(destination_dir):
        try:
            os.makedirs(destination_dir)
        except OSError:
            print("Unable to create destination dir {}".format(destination_dir))
            sys.exit(1)
    # process
    process_main(docx_file_name, destination_dir)
    pass
# ---------------------------------------------------------------------------------------------