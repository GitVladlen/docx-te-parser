import string
import re

_global_te_id = None
_global_texts = []

# MAIN =================================================
# ======================================================
def parse_encounter_nodes(encounter_nodes):
    global _global_te_id
    global _global_texts

    _global_te_id = get_id_from_nodes(encounter_nodes)
    _global_texts = []

    root = Root(None)
    cur_node = root
    for tag, value in encounter_nodes:
        try:
            cur_node = cur_node.push(tag, value.strip())
        except NoHandleTagException as ex:
            print(ex)
            pass

    script_text = str(root)

    # DEBUG
    # print("=========================", _global_te_id)
    # print(script_text)
    # print("=========================")

    return _global_te_id, script_text, _global_texts
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
                raise NoHandleTagException("Parent is None", tag, value)
            result_node = self.parent.push(tag, value)
            if result_node is None:
                raise NoHandleTagException("Parent can`t handle tag", tag, value)
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
            global _global_te_id
            global _global_texts

            TextID = "ID_TE_{ID}_{index}".format(ID=_global_te_id, index=len(_global_texts))
            TextValue = value
            _global_texts.append(dict(key=TextID, value=TextValue))

            complex_node.params["Text"] = TextID
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
            Conditions=[],
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

    def _onCheckConditions(self, context):
        {Conditions:if:conditions = [{Conditions:repeat:{{item}}}]
        if not all(conditions):
            return False}
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
            self.addValue("InitConditions", value, InitConditionWorld)
        elif tag == "Stages":
            self.addValue("InitConditions", value, InitConditionStages)
        elif tag == "Occurrence":
            self.addValue("InitConditions", value, InitConditionOccurrence)
        elif tag == "Conditions":
            return self.addComplex("Conditions", value, Conditions)
        elif tag == "Dialog":
            return self.addComplex("Dialog", value, Dialog)
        else:
            return None
        return self


# ======================================================
class Dialog(ComplexNode):
    def _onInit(self):
        self.params.update(dict(
            Text = None,
            Options = [],
            Outcome = None,
        ))    
        self.to_str_format = """
        dialog.text = "{Text}"
        {Outcome:if:{Outcome}}{Options:if:{Options:repeat:{{item}}}}"""
        
    def _push(self, tag, value):
        if tag == "Option":
            return self.addComplex("Options", value, Option)
        elif tag == "Outcome":
            return self.addComplex("Outcome", value, DialogOutcome)
        else:
            return None
        return self


# =====================================================
class Option(ComplexNode):
    def _onInit(self):
        self.params.update(dict(
            Text = None,
            Conditions = [],
            Outcomes = [],
        ))
        self.to_str_format = """
        option = dialog.option()
        option.text = "{Text}"
        {Outcomes:repeat:{{item}}}
        {Conditions:if:option_conditions = [{Conditions:repeat:{{item}}}]
        if not all(option_conditions):
            dialog.options.remove(option)}
        """

    def _push(self, tag, value):
        if tag == "Conditions":
            return self.addComplex("Conditions", value, Conditions)
        if tag == "Outcome":
            return self.addComplex("Outcomes", value, OptionOutcome)
        else:
            return None
        return self


# =====================================================
class OutcomeNode(ComplexNode):
    def _onInit(self):
        self.params.update(dict(
            Text = None,
            Conditions = [],
            Gips = None,
        ))

    def _push(self, tag, value):
        if tag == "Gips":
            self.addValue("Gips", value, OutcomeGips)
        elif tag == "Conditions":
            return self.addComplex("Conditions", value, Conditions)
        else:
            return None
        return self


# =====================================================
class OptionOutcome(OutcomeNode):
    def _onInit(self):
        super(OptionOutcome, self)._onInit()
        self.to_str_format = """
        outcome = option.outcome()
        {Text:if:outcome.text = "{Text}"}{Gips:if:{Gips}}

        {Conditions:if:outcome_conditions = [{Conditions:repeat:{{item}}}]
        if not all(outcome_conditions):
            option.outcomes.remove(outcome)}
        """


# =====================================================
class DialogOutcome(OutcomeNode):
    def _onInit(self):
        super(DialogOutcome, self)._onInit()
        self.to_str_format = """
        outcome = dialog.outcome()
        {Text:if:outcome.text = "{Text}"}{Gips:if:{Gips}}

        {Conditions:if:outcome_conditions = [{Conditions:repeat:{{item}}}]
        if not all(outcome_conditions):
            dialog.outcome = None}
        """


# OUTCOMES ============================================
class OutcomeGips(ValueNode):
    def _onInit(self):
        self.params["Value"] = None
        self.to_str_format = "\n        outcome.gips = {Value}\n"

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


# INIT CONDITIONS ======================================
class InitConditionWorld(ValueNode):
    def _onInit(self):
        self.params["World"] = None
        self.to_str_format = "self.world = \"{World}\""

    def _parse(self, value):
        self.params["World"] = value
        return True


# ======================================================
class InitConditionStages(ValueNode):
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


# =======================================================
class InitConditionOccurrence(ValueNode):
    def _onInit(self):
        self.params.update(dict(
            Value=None,
        ))
        self.to_str_format = "self.occurence = {Value}"

    def _parse(self, value):
        match = re.match(r'(Reccuring|Resets|Until completed|Once only)', str(value))

        if not match:
            return False
        bones = {
            "Reccuring": "self.OCCURRENCE_RECCURING",
            "Resets": "self.OCCURRENCE_RESETS",
            "Until completed": "self.OCCURRENCE_UNTIL_COMPLETED",
            "Once only": "self.OCCURRENCE_ONCE_ONLY",
        }
        self.params["Value"] = bones[match.group(1)]

        return True
        pass


# CONDITIONS ============================================
class Conditions(ComplexNode):
    def _onInit(self):
        self.params.update(dict(
            Mech1 = None,
            Mech2 = None,
        ))
        self.to_str_format = "{Mech1:if:{Mech1}}{Mech2:if:{Mech2}}"

    def _push(self, tag, value):
        if tag == "Mech1":
            self.addValue("Mech1", value, ConditionMech1)
        elif tag == "Mech2":
            self.addValue("Mech2", value, ConditionMech2)
        else:
            return None
        return self


# ======================================================
class ConditionMechNode(ValueNode):
    def _onInit(self):
        self.params.update(dict(
            Items = [],
            MechObjectName = None,
        ))
        self.to_str_format = "{Items:repeat:\n            context.{MechObjectName}.{{item}},}"

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
class ConditionMech1(ConditionMechNode):
    def _onInit(self):
        super(ConditionMech1, self)._onInit()
        self.params["MechObjectName"] = "mech1"


# ======================================================
class ConditionMech2(ConditionMechNode):
    def _onInit(self):
        super(ConditionMech2, self)._onInit()
        self.params["MechObjectName"] = "mech2"


# UTILS ===============================================
# =====================================================
def get_id_from_nodes(nodes):
    for tag, value in nodes:
        if tag == "ID":
            return value.strip()
    return None

class NoHandleTagException(Exception):
    def __init__(self, message, tag, value):
        self.message = message
        self.tag = tag
        self.value = value

    def __str__(self):
        return "[TAG EXCEPTION] MESSAGE='{ex.message}' TAG='{ex.tag}' VALUE='{ex.value}'".format(ex=self)


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