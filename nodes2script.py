import string
import re


# MAIN =================================================
# ======================================================
def parse_encounter_nodes(encounter_nodes):
    root = Root(None)

    cur_node = root
    for tag, value in encounter_nodes:
        try:
            cur_node = cur_node.push(tag, value.strip())
        except NoHandleTagException as ex:
            print(ex)

    script_text = root.getScript()
    te_id = root.getID()
    text_resources = root.getTexts()

    print("=========================", te_id)
    print(script_text)
    print("=========================")

    return te_id, script_text, text_resources
    pass


# BASE CLASSES =========================================
# ======================================================
class Node(object):
    def __init__(self, parent):
        self.parent = parent
        self.params = dict()
        self.to_str_format = ""

        if self.onInit() is False:
            print ("Init failed")

    def onInit(self):
        return self._onInit()

    def _onInit(self):
        return False

    def __str__(self):
        sf = SuperFormatter()
        return sf.format(self.to_str_format, **self.params)


# ======================================================
class ValueNode(Node):
    def parse(self, value):
        return self._parse(value)

    def _parse(self, value):        
        return False


# ======================================================
class ComplexNode(Node): 
    def push(self, tag, value):
        result_node = self._push(tag, value)
        if result_node is None:
            if self.parent is None:
                raise NoHandleTagException("Parent is None", type(self), tag, value)
            result_node = self.parent.push(tag, value)
            if result_node is None:
                raise NoHandleTagException("Parent can`t handle tag", type(self), tag, value)
        return result_node

    def _push(self, tag, value):
        return None

    def _add_param(self, key, value):
        if isinstance(self.params[key], list):
            self.params[key].append(value)
        else:
            self.params[key] = value

    def addValue(self, key, value, type_=None):
        if type_ is not None:
            node = type_(self)
            if node.parse(value):
                value = node
        self._add_param(key, value)
        

    def addComplex(self, key, value, type_):
        complex_node = type_(self)
        if value:
            complex_node.params["Text"] = value
        self._add_param(key, complex_node)
        return complex_node


# DERIVATIVE CLASSES ===================================
# ======================================================
class Root(ComplexNode):
    def _onInit(self):
        self.params.update(dict(
            ID=None,
            Name=None,
            InitConditions=[],
            Conditions=None,
            Dialog=None,
        ))
# format string ---------------------------------------
        self.to_str_format = """from Game.TextEncounters.TextEncounter import TextEncounter


class TextEncounter{ID}(TextEncounter):
    def __init__(self):
        super(TextEncounter{ID}, self).__init__()
        self.id = "{ID}"
        self.name = "{Name}"
{InitConditions:repeat:\n        {{item}}}
        pass

    def _onCheckConditions(self, context):{Conditions}
        return True
        pass

    def _onGenerate(self, context, dialog):{Dialog}
        pass
    pass
"""
# -----------------------------------------------------

    def _push(self, tag, value):
        if tag == "ID":
            self.addValue("ID", value)
        elif tag == "Name":
            self.addValue("Name", value)
        elif tag == "World":
            self.addValue("InitConditions", value, World)
        elif tag == "Stages":
            self.addValue("InitConditions", value, Stages)
        elif tag == "Conditions":
            return self.addComplex("Conditions", value, RootConditions)
        elif tag == "Dialog":
            return self.addComplex("Dialog", value, Dialog)
        else:
            return None
        return self

    def getScript(self):
        return str(self)

    def getID(self):
        return self.params.get("ID")

    def getTexts(self):
        return []


# ======================================================
class World(ValueNode):
    def _onInit(self):
        self.params["World"] = None
        self.to_str_format = "self.world = \"{World}\""

    def _parse(self, value):
        self.params["World"] = value
        return True

# ======================================================
class Stages(ValueNode):
    def _onInit(self):
        self.params.update(dict(
            From=None,
            To=None,
        ))
        self.to_str_format = "self.stages = range({From}, {To})"

    def _parse(self, value):
        match_two_digits = re.match(r'\s*(\d+)\s+(\d+)\s*', str(value))
        match_one_digit = re.match(r'\s*(\d+)\s*', str(value))

        from_level, to_level = None, None
        if match_two_digits:
            from_level, to_level = map(int, match_two_digits.groups())
        elif match_one_digit:
            from_level = to_level = int(match_one_digit.group())

        if from_level < 0 or to_level < 0:
            print("Invalid Levels") 
            return False

        self.params["From"] = from_level
        self.params["To"] = to_level + 1
        return True
        pass


# ======================================================
class RootConditions(ComplexNode):
    def _onInit(self):
        self.params.update(dict(
            Mech1 = None,
        ))
# format string ----------------------------------------
        self.to_str_format = """
{Mech1}
"""
# ------------------------------------------------------
    def _push(self, tag, value):
        if tag == "Mech1":
            self.addValue("Mech1", value, ConditionMech1)
        else:
            return None
        return self

# ======================================================
class ConditionMech1(ValueNode):
    def _onInit(self):
        self.params.update(dict(
            Items = [],
        ))
# format string ----------------------------------------
        self.to_str_format = """
{Items:repeat:        if not context.mech1.{{item}}:\n            return False\n}
"""
# ------------------------------------------------------
    def _parse(self, value):
        matches = re.findall(r'(HP|TAP|HE)\s*(>|>=|<|<=|==|!=)\s*(\d+)', value)
        if not matches:
            return False
        format_string = "{stat} {operator} {number}"
        for stat, operator, number in matches:
            bones = dict(
                stat=stat.lower(),
                operator=operator,
                number=number,
            )
            self.params["Items"].append(format_string.format(**bones))
        return True

# ======================================================
class Dialog(ComplexNode):
    def _onInit(self):
        self.params.update(dict(
            Text = None,
            Options = [],
        ))
# format string ---------------------------------------        
        self.to_str_format = """
        dialog.text = "{Text}"
{Options:repeat:{{item}}}
"""
# -----------------------------------------------------
        
    def _push(self, tag, value):
        if tag == "Option":
            return self.addComplex("Options", value, Option)
        else:
            return None
        return self


# =====================================================
class Option(ComplexNode):
    def _onInit(self):
        self.params.update(dict(
            Text = None,
            Outcomes = [],
        ))
# format string ---------------------------------------
        self.to_str_format = """
        option = dialog.option()
        option.text = "{Text}"
{Outcomes:repeat:{{item}}}
"""
# -----------------------------------------------------

    def _push(self, tag, value):
        if tag == "Outcome":
            return self.addComplex("Outcomes", value, Outcome)
        else:
            return None
        return self


# =====================================================
class Outcome(ComplexNode):
    def _onInit(self):
        self.params.update(dict(
            Text = None,
            Gips = None,
        ))
# format string ---------------------------------------
        self.to_str_format = """
        outcome = option.outcome()
        outcome.text = "{Text}"
        {Gips}
"""
# -----------------------------------------------------

    def _push(self, tag, value):
        if tag == "Gips":
            self.addValue("Gips", value, OutcomeGips)
        else:
            return None
        return self


# =====================================================
class OutcomeGips(ValueNode):
    def _onInit(self):
        self.params["Value"] = None
        self.to_str_format = "outcome.gips = {Value}"

    def _parse(self, value):
        rand_match = re.match(r'rand\s*[+-]?(\d+)\s+[+-]?(\d+)', str(value))
        int_match = re.match(r'[+-]?(\d+)', str(value))
        if rand_match:
            from_, to_ = rand_match.groups()
            self.params["Value"] = "self.rand({}, {})".format(from_, to_)
        elif int_match:
            int_value = int_match.group(1)
            self.params["Value"] = int_value
        else:
            return False
        return True


# UTILS ===============================================
# =====================================================
class NoHandleTagException(Exception):
    def __init__(self, message, type_, tag, value):
        self.message = message
        self.type_ = type_
        self.tag = tag
        self.value = value

    def __str__(self):
        return "TYPE='{ex.type_}' MESSAGE='{ex.message}' TAG='{ex.tag}' VALUE='{ex.value}'".format(ex=self)


# ======================================================
class SuperFormatter(string.Formatter):
    """World's simplest Template engine."""

    def format_field(self, value, spec):
        if spec.startswith('repeat'):
            template = spec.partition(':')[-1]
            if type(value) is dict:
                value = value.items()
            return ''.join([template.format(item=item) for item in value])
        elif spec == 'call':
            return value()
        elif spec.startswith('if'):
            return (value and spec.partition(':')[-1]) or ''
        else:
            return super(SuperFormatter, self).format_field(value, spec)
# ====================================================