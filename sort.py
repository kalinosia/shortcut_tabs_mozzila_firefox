import os
import sys
import winshell
import win32com.client
import string

shell = win32com.client.Dispatch('WScript.Shell')
userDesktop = winshell.desktop()
main_path = userDesktop + "\\KARTY PRZEGLĄDARKA"

# Check command-line arguments
if len(sys.argv) != 2:
    sys.exit("Usage: python sort.py name_file.txt")

path ="C:\\Users\\wavy5\\Desktop\\KARTY PRZEGLĄDARKA\\"
file_to_open=path+sys.argv[1]

def what_category(link):
	categories = {"youtube\\" : [ "youtube"],
		"shoppping\\" : ["allegro","aliexpress", "ceneo", "empik", "lidl", "sklep", "jysk", "ikea"],
		"shoppping\\beauty\\" : ["fragrantica","perfumy" , "makeup", "ezebra", "ekobieca","sinsay","drogerienatura", "drogeria", "apteka", "pigment"],
		"shoppping\\other\\" : ["pinsola", "coloeland", "zalando",  "sznur", "stradivarius", "zara", "hm","reserved", ],

		"programming\\": ["geek","instructables","skillshot", "github", "cpp", "dev", "hack", "stackoverflow"],
		"programming\\web\\" : ["css", "template", "bootstrap" ],
		"programming\\elektronics\\" : ["forbot","openeeg", "raspberry", "arduino", "skillshot"],
		"learn\\":["edx", "course", "learn", "english", "ai", "marki", "learn", "artificial", "education", "experiments", "bigdata", "wikipedia", "academy"],
	
		"our_home\\": ["stodola"],
		"media\\" : ["mail", "docs", "facebook", "canva", "filmweb", "biedronka", "bank"],
		"mems\\" : ["jebzdzidy", "jbzd", "kwejk", "9gag"] }

	for category in categories:
		for word in category:
			if word in link:
				return category

	return "other\\"


def safe_link(target, category):
	characters=["<",">",":","/","\"","\\","|","?","*","+",".","=","https"]
	# name =  target.replaceAll("<>:\"/\\|?*", "")
	name=target
	for ch in characters:
		name = name.replace(ch,"")

	name = name[:127] # trimm to len 260\255\127
	print(len(name))
	path_all = path + category  + name + r'.url' 
	#target = r"https://www.codespeedy.com/create-the-shortcut-of-any-file-in-windows-using-python/"
	print("###################")
	print(path_all)
	shortcut = shell.CreateShortcut(path_all)
	shortcut.TargetPath = target
	shortcut.save() 





with open(file_to_open, encoding="utf8") as f:
	while True:
		link = f.readline()

		if not link:
			break

		category=what_category(link)
		print(link)
		safe_link(link, category)

  



