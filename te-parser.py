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

def parse(text, file_name):
	lines = text.split("\n\n")
	lines = delete_comments(lines)
	
	result = """from TextEncounter import TextEncounter

	
class TextEncounter{ID}(TextEncounter):
	def __init__(self):
		super(TextEncounter{ID}, self).__init__()
		self.id = "{ID}"
		pass
	def _onCheckConditions(self, context):
		return True
		pass
	def _onGenerate(self, context, dialog):
		{dialog}
		pass
	pass
"""
	params = {}
	keywords = ["ID:", "Planet:"]
	
	for line in lines:
		words = line.split(" ")
		
		for index, word in enumerate(words):
			if word in keywords:
				keyword = word[:-1]
				params[keyword] = words[index + 1]
	
	params["dialog"] = "dialog.text = \"ID_TE_Text\""
	
	result = result.format(**params)
	
	with open(file_name, "w") as file:
		file.write(result)


if __name__ == '__main__':
	text = docx2txt.process("D:\\Python-Projects\\docx-to-text\\te-format.docx")
	parse(text, "D:\\Python-Projects\\docx-to-text\\result.py")
	
	