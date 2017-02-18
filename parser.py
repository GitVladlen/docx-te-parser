import sys
import os
import re
import argparse

import docx2txt

# ------------------------- USE THIS IN MENGINE -------------------------
def process(docx_file_path):
    # gain text from docx
    text = docx2txt.process(docx_file_path)
    text = text.encode('utf-8')

    # text parsing text and creating script file
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
    for nodes in nodes_by_encounters:
        # transform nodes into data
        te_id, script_text, texts = parse_data(nodes)
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
        "Planet",
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

def split_nodes_by_encounters(nodes):
    # todo: refactor to regex
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
    debug_format = """{script}
\"\"\" debug info
------------- Nodes -------------
{nodes}
------------- Texts -------------
{texts}
\"\"\"
"""
    nodes_str = "\n".join(map(lambda node: "{} = \"{}\"".format(node[0], node[1]), nodes))
    texts_str = "\n".join(map(lambda text: text_id_format.format(**text), texts))
    full_text = debug_format.format(script=script_text, nodes=nodes_str, texts=texts_str)
    return full_text
    pass

def get_id_from_nodes(nodes):
    for tag, value in nodes:
        if tag == "ID":
            return value
    return None

# -------------------------------- NODES PARSING --------------------------------
def parse_data(nodes):
    # todo: modify parsing alghoriphm

    te_id = get_id_from_nodes(nodes)

    params = dict(
        ID=te_id,
        InitParams="",
        CheckConditions="",
        GenerateDialog="",
        Texts=[],
        OptionIndex=0,
        OutcomeIndex=0,
    )

    indentation = "        "
    entity = "self"

    option_index = 0
    outcome_index = 0

    def rule_strip_space(params, key, tag, indentaion, entity, value, format_string):
        value = value.strip(" ")
        params[key] += format_string.format(
            indentation=indentation,
            entity=entity,
            value=value,
        )
        return entity
        pass

    def rule_id(params, key, tag, indentaion, entity, value, format_string):
        params["ID"] = value.strip(" ")
        return rule_strip_space(params, key, tag, indentation, entity, value, format_string)
        pass

    def rule_levels(params, key, tag, indentation, entity, value, format_string):
        match_two_digits = re.match(r'\s*(\d+)\s+(\d+)\s*', str(value))
        match_one_digit = re.match(r'\s*(\d+)\s*', str(value))

        from_level, to_level = None, None
        if match_two_digits:
            from_level, to_level = map(int, match_two_digits.groups())
        elif match_one_digit:
            from_level = to_level = int(match_one_digit.group())

        if from_level < 0 or to_level < 0:
            print("Invalid Levels in {ID}".format(**params)) 
            return entity

        value = "range({}, {})".format(from_level, to_level + 1)
        return rule_strip_space(params, key, tag, indentation, entity, value, format_string)
        pass

    def rule_conditions(params, key, tag, indentation, entity, value, format_string):
        if entity != "self":
            return "conditions"
        return entity
        pass

    def rule_mech1_conditions(params, key, tag, indentation, entity, value, format_string):
        matches = re.findall(r'(HP|TAP|HE)\s*(>|>=|<|<=|==|!=)\s*(\d+)', str(value))

        if matches:
            format_string = "\n{indentation}if not context.mech1.{stat} {operator} {number}:\n{indentation}    return False"
            for stat, operator, number in matches:
                bones = dict(
                    stat=stat.lower(),
                    indentation=indentation,
                    operator=operator,
                    number=number,
                )
                params["CheckConditions"] += format_string.format(**bones)
                pass
            pass

        return entity
        pass 

    def rule_mech1_outcome(params, key, tag, indentation, entity, value, format_string):
        params["GenerateDialog"] += "\n{indentation}# {entity}.mech1 = {value}".format(
            indentation=indentation,
            entity=entity,
            value=value,
        )
        return entity
        pass

    def rule_mech2_conditions(params, key, tag, indentation, entity, value, format_string):
        matches = re.findall(r'(HP|TAP|HE)\s*(>|>=|<|<=|==|!=)\s*(\d+)', str(value))
        if matches:
            format_string = "\n{indentation}if not context.mech1.{stat} {operator} {number}:\n{indentation}    return False"
            for stat, operator, number in matches:
                bones = dict(
                    stat=stat.lower(),
                    indentation=indentation,
                    operator=operator,
                    number=number,
                )
                params["CheckConditions"] += format_string.format(**bones)
                pass
            pass

        return entity
        pass

    def rule_mech2_outcome(params, key, tag, indentation, entity, value, format_string):
        params["GenerateDialog"] += "\n{indentation}# {entity}.mech2 = {value}".format(
            indentation=indentation,
            entity=entity,
            value=value,
        )
        return entity
        pass

    def rule_int(params, key, tag, indentation, entity, value, format_string):
        match = re.match(r'(\d+)', str(value))

        if match:
            value = int(value)
            params[key] += format_string.format(
                indentation=indentation,
                entity=entity,
                value=value,
            )
        return entity
        pass

    def rule_occurrence(params, key, tag, indentation, entity, value, format_string):

        match = re.match(r'(Reccuring|Resets|Until completed|Once only)', str(value))

        if match:
            bones = {
                "Reccuring": "self.OCCURRENCE_RECCURING",
                "Resets": "self.OCCURRENCE_RESETS",
                "Until completed": "self.OCCURRENCE_UNTIL_COMPLETED",
                "Once only": "self.OCCURRENCE_ONCE_ONLY",
            }
            value = bones[match.group(1)]
            params[key] += format_string.format(
                indentation=indentation,
                entity=entity,
                value=value,
            )

        return entity
        pass

    def rule_cargo(params, key, tag, indentation, entity, value, format_string):
        params["CheckConditions"] += "\n{indentation}# {entity}.cargo = {value}".format(
            indentation=indentation,
            entity=entity,
            value=value,
        )
        return entity
        pass

    def rule_dialog(params, key, tag, indentation, entity, value, format_string):
        entity = "dialog"
        if value:
            text_id = "ID_TE_{ID}_Dialog".format(**params)
            params["Texts"].append(dict(key=text_id, value=value))
            value = text_id
            params[key] += format_string.format(
                indentation=indentation,
                entity=entity,
                value=value,
            )
        return entity
        pass

    def rule_option(params, key, tag, indentation, entity, value, format_string):
        params[key] += "\n\n{indentation}option = dialog.option()".format(
            indentation=indentation,
            entity=entity,
            value=value,
        )
        entity = "option"

        if value:
            text_id = "ID_TE_{ID}_Option_{OptionIndex}".format(**params)
            params["Texts"].append(dict(key=text_id, value=value))
            params["OptionIndex"] += 1
            value = text_id
            params[key] += "\n{indentation}{entity}.text = \"{value}\"".format(
                indentation=indentation,
                entity=entity,
                value=value,
            )
            pass

        return entity
        pass

    def rule_outcome(params, key, tag, indentation, entity, value, format_string):
        if params["OptionIndex"] > 0:
            entity = "option"
            pass
        params[key] += "\n\n{indentation}outcome = {entity}.outcome()".format(
            indentation=indentation,
            entity=entity,
            value=value,
        )
        entity = "outcome"

        if value.strip():
            text_id = "ID_TE_{ID}_Outcome_{OutcomeIndex}".format(**params)
            params["Texts"].append(dict(key=text_id, value=value))
            params["OutcomeIndex"] += 1
            value = text_id
            params[key] += "\n{indentation}{entity}.text = \"{value}\"".format(
                indentation=indentation,
                entity=entity,
                value=value,
            )
            pass

        return entity
        pass

    def rule_gips(params, key, tag, indentation, entity, value, format_string):
        rand_match = re.match(r'rand\s*[+-]?(\d+)\s+[+-]?(\d+)', str(value))
        int_match = re.match(r'[+-]?(\d+)', str(value))
        if rand_match:
            from_, to_ = rand_match.groups()
            value = "self.rand({}, {})".format(from_, to_)
            params[key] += "\n{indentation}{entity}.gips = {value}".format(
                indentation=indentation,
                entity=entity,
                value=value,
            )
        elif int_match:
            int_value = int_match.group(1)
            value = int_value
            params[key] += "\n{indentation}{entity}.gips = {value}".format(
                indentation=indentation,
                entity=entity,
                value=value,
            )

        return entity
        pass

    def rule_items(params, key, tag, indentation, entity, value, format_string):
        params[key] += format_string.format(
            indentation=indentation,
            entity=entity,
            value="\"{}\"  # dummy".format(value),
        )
        return entity
        pass

    rules = dict(
        ID=(rule_id, "InitParams", "\n{indentation}{entity}.id = \"{value}\""),
        Name=(rule_strip_space, "InitParams", "\n{indentation}{entity}.name = \"{value}\""),
        Conditions=(rule_conditions, None, None),
        Planet=(rule_strip_space, "InitParams", "\n{indentation}{entity}.planet = \"{value}\""),
        Levels=(rule_levels, "InitParams", "\n{indentation}{entity}.levels = {value}"),
        Priority=(rule_int, "InitParams", "\n{indentation}{entity}.priority = {value}"),
        Occurrence=(rule_occurrence, "InitParams", "\n{indentation}{entity}.occurrence = {value}"),
        Frequency=(rule_int, "InitParams", "\n{indentation}{entity}.frequency = {value}"),
        Mech1=dict(
            self=(rule_mech1_conditions, None, None),
            conditions=(rule_mech1_conditions, None, None),
            outcome=(rule_mech1_outcome, None, None),
        ),
        Mech2=dict(
            self=(rule_mech2_conditions, None, None),
            conditions=(rule_mech2_conditions, None, None),
            outcome=(rule_mech2_outcome, None, None),
        ),
        Cargo=(rule_cargo, None, None),
        Dialog=(rule_dialog, "GenerateDialog", "\n{indentation}{entity}.text = \"{value}\""),
        Option=(rule_option, "GenerateDialog", None),
        Outcome=(rule_outcome, "GenerateDialog", None),
        Chance=(rule_int, "GenerateDialog", "\n{indentation}{entity}.chance = {value}"),
        Gips=(rule_gips, "GenerateDialog", "\n{indentation}{entity}.gips = {value}"),
        Items=(rule_items, "GenerateDialog", "\n{indentation}{entity}.items = {value}")
    )

    for tag, value in nodes:
        rule = rules.get(tag)
        if rule is None:
            # print "There are no rule for tag {}".format(tag)
            continue    
        if isinstance(rule, dict) is True:
            rule = rule.get(entity)
        if rule is None:
            #print "There are no rule for tag {} in entity {}".format(tag, entity)
            continue
        action, key, format_string = rule
        entity = action(params, key, tag, indentation, entity, value.strip(), format_string)
        pass

    script_text = script_format.format(**params)

    script_text = add_debug_text(script_text, nodes, params["Texts"])

    return params["ID"], script_text, params["Texts"]
    pass

# ---------------------------- FORMAT STRINGS ----------------------------
script_format = """from Game.TextEncounters.TextEncounter import TextEncounter


class TextEncounter{ID}(TextEncounter):
    def __init__(self):
        super(TextEncounter{ID}, self).__init__(){InitParams}
        pass

    def _onCheckConditions(self, context):{CheckConditions}
        return True
        pass

    def _onGenerate(self, context, dialog):{GenerateDialog}
        pass
    pass
"""

texts_format = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Texts>
    {Texts}
</Texts>
"""

text_id_format = "<Text Key=\"{key}\" Value=\"{value}\"/>"

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