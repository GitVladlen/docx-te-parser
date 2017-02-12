import sys
import os
import docx2txt

import argparse

def process_args():
    parser = argparse.ArgumentParser(
        description="Export Text Encounter from docx file to Mengine"
        )
    parser.add_argument("docx", help="path of the docx file")
    parser.add_argument("dest", help="path of destination directory")

    args = parser.parse_args()
        
    return args
    pass

# ---------------------------- FORMAT STRINGS ----------------------------
script_format = """from TextEncounter import TextEncounter


class TextEncounter{ID}(TextEncounter):
    def __init__(self):
        super(TextEncounter{ID}, self).__init__()
        self.id = "{ID}"
        self.name = "{Name}"
{LineConditions}
        pass

    def _onCheckConditions(self, context):{ContextConditions}
        return True
        pass

    def _onGenerate(self, context, dialog):{Dialog}
        pass
    pass
"""

texts_format = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Texts>
    {Texts}
</Texts>
"""

text_id_format = "<Text Key=\"{key}\" Value=\"{value}\"/>"

# -------------------------------- UTILS --------------------------------
def delete_comments(lines):
    """Clear text from comments like //comment"""
    result = []
    for line in lines:
        # delete comments from line
        comment_index = line.find("//")
        if comment_index is not -1:
            line = line[:comment_index]
        # ignore empty line
        if not line:
            continue
        # add to result
        result.append(line)
    return result

def get_tag(word):
    """Check word for tag"""
    if word.count(":") != 1:
        return None
    if word.endswith(":") is False:
        return None
    return word[:-1]

def parse_nodes(text):
    """Parsing nodes from text
    node = tag, value
    """
    lines = text.split("\n\n")
    lines = delete_comments(lines)

    word_delimiter = " "
    args = []
    last_tag = None

    nodes = []

    for line in lines:
        words = line.split(word_delimiter)
        for word in words:
            tag = get_tag(word)
            if tag is None:
                args.append(word)
                continue

            if last_tag is not None:
                value = " ".join(args)
                value = value.replace(" [br] ", "&#10;")
                node = last_tag, value
                nodes.append(node)

                args = []
                last_tag = tag
            else:
                last_tag = tag
            pass
        pass

    return nodes
    pass

def write_script(script_text, dir_path, te_id):
    """Write python script"""
    # do full file path
    file_name = "TextEncounter{ID}.py".format(ID=te_id)
    path = os.path.join(dir_path, file_name)
    print path
    # write
    with open(path, "w") as f:
        f.write(script_text)
        pass
    pass

def write_texts(all_texts, dir_path):
    file_name = "Texts.xml"
    path = os.path.join(dir_path, file_name)
    print path
    all_texts_str_list = []
    for texts in all_texts:
        texts_str = "\n\t".join(map(lambda text: text_id_format.format(**text), texts))
        all_texts_str_list.append(texts_str)
    all_texts_str = "\n\n\t".join(all_texts_str_list)

    texts_to_write = texts_format.format(Texts=all_texts_str)
    with open(path, "w") as f:
        f.write(texts_to_write)
        pass
    pass

def add_debug_text(script_text, nodes, texts):
    """Add debug text as multi line comment to script text"""
    debug_format = """{script}
\"\"\" debug info
------------- Nodes -------------
{nodes}
------------- Texts -------------
{texts}
\"\"\"
"""
    nodes_str = "\n".join(map(lambda (t, v): "{} = \"{}\"".format(t, v), nodes))
    texts_str = "\n".join(map(lambda text: text_id_format.format(**text), texts))
    full_text = debug_format.format(script=script_text, nodes=nodes_str, texts=texts_str)
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

    # print "".join(map(lambda (index, nodes): "{}. {}\n\n".format(index, nodes), enumerate(result)))
    return result
    pass

def parse_data(nodes):
    params = dict(
        ID="000",
        Name="NoName",
        LineConditions="",
        ContextConditions="",
        Dialog="",
    )

    indentation = "        "
    texts = []

    option_index = 0
    outcome_index = 0
    entity = None

    for tag, value in nodes:
        if tag == "Conditions":
            entity = "conditions"
        elif tag == "ID":
            params["ID"] = value
        elif tag == "Name":
            params["Name"] = value
        elif tag == "Planet":
            params["LineConditions"] += "\n{}self.planet = \"{}\"".format(indentation, value)
        elif tag == "Levels":
            lvl_from, lvl_to = map(int, value.split())
            params["LineConditions"] += "\n{}self.levels = range({}, {})".format(indentation, lvl_from, lvl_to + 1)
        elif tag == "Priority":
            params["LineConditions"] += "\n{}self.priority = {}".format(indentation, int(value))
        elif tag == "Occurrence":
            params["LineConditions"] += "\n{}self.occurrence = \"{}\"".format(indentation, value)
            pass
        elif tag == "Frequency":
            params["LineConditions"] += "\n{}self.frequency = {}".format(indentation, int(value))
            pass
        elif tag == "Mech1":
            if entity is "conditions":
                params["ContextConditions"] += "\n{}if context.mech1.dummy() is False:  # {}".format(indentation, value)
                params["ContextConditions"] += "\n{}    return False\n".format(indentation)
            pass
        elif tag == "Mech2":
            pass
        elif tag == "Cargo":
            params["ContextConditions"] += "\n{}if context.cargo.dummy() is False:  # {}".format(indentation, value)
            params["ContextConditions"] += "\n{}    return False\n".format(indentation)
        elif tag == "Dialog":
            if value:
                text = dict(
                    key="ID_TE_{}_Dialog".format(params["ID"]),
                    value=value
                )
                texts.append(text)
                params["Dialog"] += "\n{}dialog.text = \"{}\"".format(indentation, text["key"])
        elif tag == "Option":
            entity = "option"
            params["Dialog"] += "\n\n{}option = dialog.option()".format(indentation)
            if value:
                option_index += 1
                text = dict(
                    key="ID_TE_{}_Option_{}".format(params["ID"], option_index),
                    value=value
                )
                texts.append(text)
                params["Dialog"] += "\n{}{}.text = \"{}\"".format(indentation, entity, text["key"])
        elif tag == "Outcome":
            entity = "outcome"
            params["Dialog"] += "\n\n{}outcome = dialog.outcome()".format(indentation)
            if value:
                outcome_index += 1
                text = dict(
                    key="ID_TE_{}_Outcome_{}".format(params["ID"], option_index),
                    value=value
                )
                texts.append(text)
                params["Dialog"] += "\n{}{}.text = \"{}\"".format(indentation, entity, text["key"])
        elif tag == "Gips":
            params["Dialog"] += "\n{}{}.gips = \"{}\"".format(indentation, entity, value)
        elif tag == "Chance":
            params["Dialog"] += "\n{}{}.chance = \"{}\"".format(indentation, entity, int(value))
        elif tag == "Chance":
            params["Dialog"] += "\n{}{}.chance = \"{}\"".format(indentation, entity, int(value))
        elif tag == "Combat":
            params["Dialog"] += "\n\n{}combat = {}.option()".format(indentation, entity)
            entity = "combat"
        elif tag == "Enemy1":
            params["Dialog"] += "\n{}{}.enemy1 = \"{}\"".format(indentation, entity, value)
        elif tag == "LoadTE":
            params["Dialog"] += "\n{}{}.loadTE = \"{}\"".format(indentation, entity, value.strip(" "))
        pass

    script_text = script_format.format(**params)
    # add debug text
    script_text = add_debug_text(script_text, nodes, texts)
    return params["ID"], script_text, texts
    pass

def process_text(text, dir_path):
    # get nodes from text
    all_nodes = parse_nodes(text)
    nodes_by_encounters = split_nodes_by_encounters(all_nodes)
    # transform nodes to files
    all_texts = []
    for nodes in nodes_by_encounters:
        te_id, script_text, texts = parse_data(nodes)
        # accumulate all texts
        all_texts.append(texts)
        # write python script
        write_script(script_text, dir_path, te_id)
        pass
    # write xml file with text resources
    write_texts(all_texts, dir_path)
    pass

# ----------------------- USE THIS -----------------------
def process_docx(docx_file_path, destination_dir):
    # gain text from docx
    text = docx2txt.process(docx_file_path)
    text = text.encode('utf-8')

    # write text to txt file, for debug
    result_txt_file_name = os.path.join(destination_dir, "docx-to-text-result.txt")
    with open(result_txt_file_name, "w") as file:
        file.write(text)

    # text parsing text and creating script file
    process_text(text, destination_dir)
    pass
# --------------------------------------------------------
if __name__ == '__main__':
    # setup default file names
    cur_dir_path = os.path.dirname(__file__)
    destination_dir = os.path.join(cur_dir_path, "debug/")
    docx_file_name = os.path.join(destination_dir, "te-format.docx")

    # todo: fix process args
    # check console args
    # args = process_args()
    # print "Command line arguments: {}".format(args)
    # if args.docx is not None and args.dest is None:
    #     if not os.path.exists(args.docx):
    #         print("File {} dose not exist.".format(args.docx))
    #         sys.exit(1)
    #         pass
    #     if not os.path.exists(args.dest):
    #         try:
    #             os.makedirs(args.dest)
    #         except OSError:
    #             print("Unable to create destination dir {}".format(args.dest))
    #             sys.exit(1)
    #         pass
        
    #     docx_file_name = args.docx
    #     destination_dir = args.dest
    #     pass
    # execute

    process_docx(docx_file_name, destination_dir)
    pass
