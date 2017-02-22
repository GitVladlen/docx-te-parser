import re
import docx2txt

from .nodes2script import parse_encounter_nodes

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
        te_id, script_text, texts = parse_encounter_nodes(encounter_nodes)
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