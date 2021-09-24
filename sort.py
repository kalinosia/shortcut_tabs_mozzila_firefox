import os
import sys
import winshell
import win32com.client
import string


shell = win32com.client.Dispatch('WScript.Shell')
userDesktop = winshell.desktop()
main_path = (os.path.join(userDesktop, "KARTY PRZEGLĄDARKA"))
print(os.path.abspath(os.getcwd()))

def main():
	# Check command-line arguments
	if len(sys.argv) != 2:
	    sys.exit("Usage: python sort.py name_file.txt")
	if not os.path.exists(os.path.abspath(sys.argv[1])):
		print("Plik z danymi (.txt) nie istnieje lub zła nazwa!")
	else:
		try:
	  
			file_to_open= os.path.abspath(sys.argv[1])
			#file_to_open=os.path.join(main_path, sys.argv[1])

			with open(file_to_open, encoding="utf8") as f:
				i=0
				while True:
					i+=1
					link = f.readline()

					if not link:
						break
					print(i)

					category=what_category(link)
					print(link, "\n", category, "\n\n")
					#safe_link(link, category)
		except:
			print("Plik nie istnieje lub zapisany w złym miejscu!\n Problem z folderami")



def what_category(link):
	categories = {"youtube" : ["youtube"],
	"shopping" : ["allegro","aliexpress", "ceneo", "empik", "lidl", "sklep", "jysk", "ikea", "biedronka"],
	"shopping beauty" : ["fragrantica","perfumy" , "makeup", "ezebra", "ekobieca","drogerienatura", "drogeria", "apteka", "pigment"],
	"shopping other" : ["pinsola", "zalando",  "sznur", "stradivarius", "zara", "hm","reserved","sinsay"],

	"programming": ["geek","instructables","skillshot", "github", "cpp", "dev", "hack", "stackoverflow", "python"],
	"programming web" : ["css", "template", "bootstrap" ],
	"programming electronics" : ["forbot","openeeg", "raspberry", "arduino", "skillshot"],

	"learn":["programming", "edx", "course", "learn", "english", "ai", "marki", "learn", "artificial", "education", "experiments", "bigdata", "wikipedia", "academy", "history"],
	
	"our_home": ["stodola"],

	"media" : ["mail",  "facebook", "filmweb", "bank", "hbo", "amazon","netflix"],

	"projects": ["canva", "docs", "colorland"],

	"mems" : ["jebzdzidy", "jbzd", "kwejk", "9gag"]}

	for category in categories:
		for word in categories[category]:
			if word in link:
				return category

	return "other"


def safe_link(target, category):
	
	characters=["<",">",":","/","\"","\\","|","?","*","+","=","https", " "]
	# name =  target.replaceAll("<>:\"/\\|?*", "")
	
	name=target
	for ch in characters:
		name = name.replace(ch,"")

	if len(name)>=127:
		name = name[:127] # trimm to len 260\255\127
	name = name.replace(" ","")


	if " " in category:
		category=category.split()
		category=os.path.join(category[0], category[1])

	name_link=name.strip()+".url"

	path_all = (os.path.join(main_path, category, name_link))
	#target = r"https://www.codespeedy.com/create-the-shortcut-of-any-file-in-windows-using-python/"

	shortcut = shell.CreateShortCut(path_all)
	shortcut.TargetPath = target
	shortcut.save() 

if __name__ == "__main__":
    main()




  



