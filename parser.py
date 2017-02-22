import sys
import os
import re
import argparse

import docx2txt
import nodes2script

# ------------------------- USE THIS IN MENGINE -------------------------
def process(docx_file_path):
    # gain text from docx
    text = docx2txt.process(docx_file_path)
    text = text.encode('utf-8')

    # gain scripts and text resources from docx text
    return parse_text(text)
    pass

# -------------------------------- UTILS --------------------------------
def parse_text(text):
    # get nodes from text
    all_nodes = parse_nodes(text)
    nodes_by_encounters = split_nodes_by_encounters(all_nodes)
    # transform nodes to files
    scripts = []
    all_texts = []
    for encounter_nodes in nodes_by_encounters:
        # transform nodes into data
        te_id, script_text, texts = nodes2script.parse_encounter_nodes(encounter_nodes)
        # DEBUG
        script_text = add_debug_text(script_text, encounter_nodes, texts)
        # accumulate all texts
        all_texts.append(texts)
        # accumulate all scripts
        scripts.append((te_id, script_text))
        pass
    texts = format_texts(all_texts)
    return scripts, texts
    pass

def parse_nodes(text):
    """Parsing nodes from text
    node = tag, value
    """
    tag_list = [
        "ID",
        "Name",
        "Conditions",
        "World",
        "Stages",
        "Stages",
        "Priority",
        "Occurrence",
        "Frequency",
        "Mech1",
        "Mech2",
        "Cargo",
        "Dialog",
        "Option",
        "Outcome",
        "Chance",
        "Gips",
        "Items",
        "Combat",
        "Enemy1",
        "Enemy2",
        "Scrap",
        "LoadTE",
    ]
    # gain text without comments and joined by space
    text = delete_comments_from_text(text)

    text_bytes = text.encode()

    all_tags_joined = "|".join(tag_list)
    # find tags
    tags_regex = r'({}):'.format(all_tags_joined)
    tags = re.findall(tags_regex, text_bytes.decode())
    # match values
    values_regex = r'{}'.format(":(.*)".join(tags) + ":(.*)")
    value_matches = re.match(values_regex, text_bytes.decode())
    # gain values from values matches
    values = []
    if value_matches:
        values = value_matches.groups()
    # pack nodes
    nodes = list(zip(tags, values))
    return nodes
    pass

def delete_comments(lines):
    """Clear text from comments like //comment"""
    result = []
    for line in lines:
        # delete comments from line
        comment_index = line.find(b"//")
        if comment_index is not -1:
            line = line[:comment_index]
        # ignore empty line
        if not line:
            continue
        # add to result
        result.append(line.decode())
    return result

def delete_comments_from_text(text):
    lines = text.split(b'\n\n')
    result = delete_comments(lines)

    return ' '.join(result)
    pass

def format_texts(all_texts):
    all_texts_str_list = []
    for texts in all_texts:
        texts_str = "\n\t".join(map(lambda text: text_id_format.format(**text), texts))
        all_texts_str_list.append(texts_str)
    all_texts_str = "\n\n\t".join(all_texts_str_list)

    texts_to_write = texts_format.format(Texts=all_texts_str)
    return texts_to_write
    pass

def add_debug_text(script_text, nodes, texts):
    """Add debug text as multi line comment to script text"""
    nodes_str = "\n".join(map(lambda node: "{} = \"{}\"".format(node[0], node[1]), nodes))
    texts_str = "\n".join(map(lambda text: text_id_format.format(**text), texts))
    full_text = debug_text_format.format(script=script_text, nodes=nodes_str, texts=texts_str)
    return full_text
    pass

def split_nodes_by_encounters(nodes):
    result = []
    cur_nodes = []
    for node in nodes:
        tag, _ = node
        
        if tag == "ID":
            if cur_nodes:
                result.append(cur_nodes)
            cur_nodes = []

        cur_nodes.append(node)

    if cur_nodes:
        result.append(cur_nodes)

    return result
    pass

# ---------------------------- FORMAT STRINGS ----------------------------
texts_format = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Texts>
    {Texts}
</Texts>
"""

text_id_format = "<Text Key=\"{key}\" Value=\"{value}\"/>"

# DEBUG
debug_text_format = """{script}
\"\"\" debug info
------------- Nodes -------------
{nodes}
------------- Texts -------------
{texts}
\"\"\"
"""

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