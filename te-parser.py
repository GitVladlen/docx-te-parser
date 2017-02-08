import docx2txt

def delete_comments(lines):
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

def get_value(keyword, line):
	start_index = line.find(keyword) + len(keyword) + 2
	part_line = line[start_index:].split(" ")
	
	words = []
	
	for word in part_line:
		if ":" in word:
			break
		words.append(word)
	
	value = " ".join(words)

	
	return value
	pass
	
def occurrence_check(value):
	if value == "Resets":
		return "self.OCCURRENCE_RESETS"
	pass
	
def levels_check(value):
	values = [int(_) for _ in value.split(" ")]
	return "range({}, {})".format(values[0], values[1] + 1)
	
def parse(text, file_name):
	text = text.encode('utf-8')
	lines = text.split("\n\n")
	lines = delete_comments(lines)
	
	result = """from TextEncounter import TextEncounter

	
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
	keywords = dict(
		ID=str,
		Name=str,
		Priority=int,
		Occurrence=occurrence_check,
		Frequency=int,
		Planet=str,
		Levels=levels_check,
	)
	
	for line in lines:
		words = line.split(" ")
		
		for word in words:
			if ":" not in word:
				continue
			keyword = word[:-1]
			
			if keyword not in keywords:
				continue
				
			handler = keywords[keyword]
			params[keyword] = handler(get_value(keyword, line))
	
	
	result = result.format(**params)
	
	with open(file_name, "w") as file:
		file.write(result)


if __name__ == '__main__':
	text = docx2txt.process("D:\\Python-Projects\\docx-to-text\\te-format.docx")
	parse(text, "D:\\Python-Projects\\docx-to-text\\result.py")
	
	
