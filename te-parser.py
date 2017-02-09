import docx2txt

# utils for parsing
def delete_comments(lines):
    """
    Clear text from comments like //comment
    """
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
    """
    Check word for tag
    """
    if word.count(":") != 1:
        return None
    if word.endswith(":") is False:
        return None
    return word[:-1]

def parse_nodes(text):
    """
    Parsing nodes from text
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

# main function
def parse(text, file_name):
    nodes = parse_nodes(text)

    result = """
from TextEncounter import TextEncounter


class TextEncounter{ID}(TextEncounter):
    def __init__(self):
        super(TextEncounter{ID}, self).__init__()
        self.id = "{ID}"
        self.name = "{Name}"
        self.priority = {Priority}
        self.occurrence = {Occurrence}
        self.frequency = {Frequency}

        self.planet = "{Planet}"
        self.levels = {Levels}
        pass
    def _onCheckConditions(self, context):
        return True
        pass
    def _onGenerate(self, context, dialog):
        pass
    pass
"""
    params = dict(
        ID=None,
        Name=None,
        Priority=None,
        Occurrence=None,
        Frequency=None,
        Planet=None,
        Levels=None,
    )

    debug_text = ""

    # process nodes
    text_ids = ""
    text_id_format = "<Text Key=\"{id}\" Value=\"{value}\"/>\n"
    for tag, value in nodes:
        if tag in params:
            params[tag] = value
        #dummy
        if tag == "Dialog" and value:
            text_ids += text_id_format.format(id="ID_TE_Dialog", value=value)

        if tag == "Option" and value:
            text_ids += text_id_format.format(id="ID_TE_Option", value=value)

        if tag == "Outcome" and value:
            text_ids += text_id_format.format(id="ID_TE_Outcome", value=value)

        debug_text += "{} = \"{}\"\n".format(tag, value)
        pass

    result = result.format(**params)

    # dummy write files
    with open(file_name, "w") as file:
        file.write(result)
        file.write("\n\"\"\" debug info\n")
        file.write("--------- ID - Value ---------\n")
        file.write(debug_text)
        file.write("----------- Texts -----------\n")
        file.write(text_ids)
        file.write("\"\"\"")
        pass

    pass

# debug tests
if __name__ == '__main__':
    # gain text from docx
    text = docx2txt.process("D:\\Python-Projects\\docx-to-text\\te-format.docx")
    text = text.encode('utf-8')
    # write text to txt file, for debug
    with open("D:\\Python-Projects\\docx-to-text\\docx-to-text-result.txt", "w") as file:
        file.write(text)
    # text parsing text and creating script file
    parse(text, "D:\\Python-Projects\\docx-to-text\\text-to-script-result.py")
    pass
