import xlrd, getopt, sys, re, os


def enum (**enums):
	return type("Enum", (), enums)

def cmplist(l, r):
	for i in range(len(l)):
		if l[i] == r[i]:
			continue
		elif l[i] < r[i]:
			return -1
		elif l[i] > r[i]:
			return 1
	return 0

def pwd():
	path = sys.path[0]
	if os.path.isfile(path):
		path = os.path.dirname(path)
	return path

def mkdir(path):
	if not os.path.exists(path):
		os.makedirs(path)

ExportTag = enum(
	NotExport = 1,
	BothSide = 2,
	SeverSide = 3,
	ClientSide = 4,
	Language = 5,
	End = 6,
	ExportName = 7,
)

Color2ExportTag = {
	# 灰 128.128,128
	23 : ExportTag.NotExport,
	# 绿 0,255,0
	3 : ExportTag.BothSide,
	11 : ExportTag.BothSide,
	# 青 0,255,255
	7 : ExportTag.SeverSide,
	15 : ExportTag.SeverSide,
	35 : ExportTag.SeverSide,
	# 黄 255,255,0
	5 : ExportTag.ClientSide,
	13 : ExportTag.ClientSide,
	34 : ExportTag.ClientSide,
	# 红 255,0,0
	2 : ExportTag.Language,
	10 : ExportTag.Language,
	# 黑 0,0,0
	0 : ExportTag.End,
	8 : ExportTag.End,
	# 紫 255,0,255
	6 : ExportTag.ExportName,
	14 : ExportTag.ExportName,
	33 : ExportTag.ExportName,
}

def parse(f):
	workbook = xlrd.open_workbook(f, formatting_info=True)
	res = {
		"name" : {},
		"keys" : [],
		"server_keys" : [],
		"server_sheet" : [],
		"client_keys" : [],
		"client_sheet" : [],
		"language_keys" : [],
		"sheet" : [],
	}
	
	def bg(cell):
		return workbook.xf_list[cell.xf_index].background.pattern_colour_index
	def tag(cell):
		return Color2ExportTag.get(bg(cell))

	def get_name(sheet):
		for row in range(sheet.nrows):
			cur_row = row
			cell = sheet.cell(row, 0)
			if tag(cell) == ExportTag.ExportName:
				return cell.value
			if tag(cell):
				break

	def findstart(sheet):
		for row in range(sheet.nrows):
			cell = sheet.cell(row, 0)
			t = tag(cell)
			if t and t != ExportTag.ExportName:
				return row		
	
	def findend(sheet):
		for row in range(sheet.nrows):
			if tag(sheet.cell(row, 0)) == ExportTag.End:
				return row		

	def split_keys(sheet, row):
		keys = sheet.row_values(row)
		server_keys = []
		client_keys = []
		language_keys = []
		for col in range(sheet.ncols):
			t = tag(sheet.cell(row, col))
			if t == ExportTag.BothSide or t == ExportTag.SeverSide:
				server_keys.append(keys[col])
			if t == ExportTag.BothSide or t == ExportTag.Language or t == ExportTag.ClientSide:
				client_keys.append(keys[col])
			if t == ExportTag.Language:
				language_keys.append(keys[col])
		return [keys, server_keys, client_keys, language_keys]

	def do_parse(sheet, start, end):
		def split_row(row):
				row_values = sheet.row_values(row)
				ds = []
				dc = []
				for col in range(sheet.ncols):
					t = tag(sheet.cell(start - 1, col))
					if t == ExportTag.BothSide or t == ExportTag.SeverSide:
						ds.append(row_values[col])
					if t == ExportTag.BothSide or t == ExportTag.Language or t == ExportTag.ClientSide:
						dc.append(row_values[col])
				res["server_sheet"].append(ds)
				res["client_sheet"].append(dc)
				res["sheet"].append(row_values)


		for row in range(start, end):

			if tag(sheet.cell(row, 0)) == ExportTag.End:
				break

			v = sheet.cell(row, 0).value
			# print(v)
			if isinstance(v, str) and (v == "" or v[0:2] == "//"):
				continue
			# print("do_parse", row)
			split_row(row)

		


	# start 
	sheet0 = workbook.sheets()[0]
	res["name"] = get_name(sheet0)
	if not res["name"]:
		return 

	keys = split_keys(sheet0, findstart(sheet0))
	res["keys"] = keys[0]
	res["server_keys"] = keys[1]
	res["client_keys"] = keys[2]
	res["language_keys"] = keys[3]

	for sheet in workbook.sheets():
		start = findstart(sheet)
		end = findend(sheet)
		if not start:
			continue
		ks = split_keys(sheet, start)
		for i in range(len(ks)):
			if cmplist(ks[i], keys[i]) != 0:
				return
		do_parse(sheet, start + 1, end)
		# print(start, end)

	return res




tpl_define = """
module("resmng")

<<content>>
"""

tpl_server = """
module("resmng")
config<<name>> = {
<<content>>
}
"""

tpl_client = """
module("resmng")

config<<name>> = {
<<content>>
}

_config<<name>>Key = {
<<keys>>
}

_config<<name>>LanguageKey = {
<<lankeys>>
}

_mt<<name>> = {
	__index = function(t, k)
		local a = t[_config<<name>>Key[k]]
		if _config<<name>>LanguageKey[k] ~= nil then
			return configLanguage[t[_config<<name>>Key[k]]]
		else
			return t[_config<<name>>Key[k]]
		end
	end
}

for k, v in pairs(config<<name>>) do
	setmetatable(v, _mt<<name>>)
end

"""

tpl_language = """
module("resmng")

config<<name>> = {
<<content>>	
}

"""

def smart_str(v):
	if isinstance(v, float):
		if v == int(v):
			v = int(v)
	elif isinstance(v, str):
		if v == "":
			v = "nil"
	return str(v)

def tplc(tpl, args):
	# print(dir(args))
	for (k, v) in args.items():
		tpl = re.sub("<<"+k+">>",smart_str(v),tpl)
	# print(tpl)
	return tpl

def export_lua(t):
	def mk_define(name):
		rows = ""
		for idx, row in enumerate(t["sheet"]):
			cell=tplc("<<key>>=<<value>>", {"key":row[0],"value":row[1]})
			rows += tplc("<<cell>>\n", {"cell":cell})
		content = tplc(tpl_define, {"content":rows})
		return content

	def mk_server_config(name):
		lines = ""
		for rowid, row in enumerate(t["server_sheet"]):
			cell = ""
			for colid, col in enumerate(row):
				cell += tplc("<<key>>=<<value>>,", {"key":t["server_keys"][colid],"value":col})
			lines += tplc("    [<<id>>]={<<cell>>},\n",{"id":row[0],"cell":cell})
		content = tplc(tpl_server,{"name":name,"content":lines})
		return content

	def mk_client_config(name):
		lines = ""
		for rowid, row in enumerate(t["client_sheet"]):
			cell = ""
			for colid, col in enumerate(row):
				cell += tplc("<<value>>,", {"key":t["client_keys"][colid],"value":col})
			lines += tplc("    [<<id>>]={<<cell>>},\n",{"id":row[0],"cell":cell})
		keys = ""
		for idx, v in enumerate(t["client_keys"]):
			keys += tplc("    <<v>> = <<k>>,\n",{"k":idx+1,"v":v})

		lankeys = ""
		for idx, v in enumerate(t["language_keys"]):
			lankeys += tplc("    <<v>> = <<k>>,\n",{"k":idx+1,"v":v})
		content = tplc(tpl_client,{
			"name":name, 
			"content":lines,
			"keys":keys,
			"lankeys":lankeys,
		})
		return content

	def get_export_name(name):
		if name[0:6] == "config":
			return name[6:]

	name = get_export_name(t["name"])
	if not name:
		print("unexpected export name:" + t["name"])
		return


	path = pwd()
	mkdir(path+"/../_out")
	mkdir(path+"/../_out/server")
	mkdir(path+"/../_out/client")

	if t["keys"][1] == "EnumID":
		content = mk_define(name)
		f = open(path + "/../_out/server/define"+name+".lua", "w")
		f.write(content)
		f.close()
		f = open(path +"/../_out/client/define"+name+".lua", "w")
		f.write(content)
		f.close()

	if name == "Language":
		for idx, v in enumerate(t["client_keys"][1:]):
			f = open(path +"/../_out/client/config"+name+v+".lua", "w")
			rows = ""
			for rowid, row in enumerate(t["client_sheet"]):
				rows += tplc("    <<k>>=<<v>>,\n",{"k":row[0],"v":row[idx+1]})
			content = tplc(tpl_language, {"content":rows,"name":name})
			f.write(content)
			f.close()


	else:
		f = open(path +"/../_out/server/config"+name+".lua", "w")
		f.write(mk_server_config(name))
		f.close()

		
		f = open(path +"/../_out/client/config"+name+".lua", "w")
		f.write(mk_client_config(name))
		f.close()
	return True

def main():
	if len(sys.argv) == 0:
		print("no input file")
		return

	path = sys.argv[1]

	if os.path.basename(path)[0:6] != "config" or os.path.splitext(path)[1][1:] != "xls":
		print("escape "+path)
		return
	

	t = parse(path)
	if not t:
		print("!!!!!!! Parse Failed !!!!!!!"+path)
		os.system("pause")
	if export_lua(t):
		print("export success: "+path)
	else:
		print("export faield: "+path)

if __name__ == '__main__':
	main()
		
	
	
