# py -m PyInstaller --onefile python_with_jupyter_notebook_install.py # <---------- convert py to exe
# py -m PyInstaller python_with_jupyter_notebook_install.py
from tkinter import *
import sys
import win32com.client
# import pythoncom
import os
import subprocess
import string
from tkinter import filedialog
import winreg
import shlex
current_installing_path = os.path.dirname(os.path.abspath(__file__))
try:
	class popupWindow(object):
		def __init__(self, master):
			top = self.top = Toplevel(master)
			top.title("!!!!!!")
			top.geometry("500x80") #the size of the app
			top.resizable(0, 0) #Don't allow resizing in the x or y direction
			self.l = Label(top, text = "請輸入 jupyter notebook 預設開啟的資料夾路徑")
			self.l.pack()
			self.l2 = Label(top, text = "或是你輸入的路徑不存在")
			self.l2.pack()
			self.b = Button(top, text = 'Ok', command = self.destroy_popup)
			self.b.pack()

		def destroy_popup(self):
			self.top.destroy()

	class mainWindow(object):
		def __init__(self, master):
			self.master = master
			self.master.title("Python with jupyter notebook installing")
			self.master.geometry("500x400") #the size of the app
			self.master.resizable(0, 0)
			backGroundCanvas = Canvas(master, width = 500, height = 400)
			backGroundCanvas.focus_set()
			backGroundCanvas.pack()
			self.l = Label(master, text = "Packages List (可自行新增刪減)").place(x = 5, y = 5)
			# self.l.pack()
			self.e = Text(master, width = 65, height = 15)
			default_packages_list = ["numpy",
									"pandas",
									"matplotlib",
									"Pillow",
									"opencv-python",
									"ipywidgets",
									"IPython",
									"keras",
									"sklearn",
									"tensorflow",
									]
			for default_package_name in default_packages_list:
				self.e.insert(END, default_package_name + "\n")
			self.e.place(x = 18, y = 30)
			# self.e.pack()
			self.l2 = Label(master, text = "選擇 jupyter notebook 預設開啟的資料夾路徑")
			self.l2.place(x = 5, y = 240)
			# self.l2.pack()

			self.path = StringVar()
			self.l3 = Entry(master, textvariable = self.path, width = 65)
			self.l3.place(x = 18, y = 265) 
			# self.l3.pack()
			self.b2 = Button(master, text = "Browse", command = self.browse_button)
			self.b2.place(x = 430, y = 260) 
			# self.b2.pack()
			self.b = Button(master, text = 'Next', command = self.cleanup, height = 3, width = 15)
			self.b.place(x = 190, y = 320)
			# self.b.pack()

		def entryValue(self):
			return self.packages, self.path.get()

		def cleanup(self):
			self.packages = self.e.get("1.0", END)
			if os.path.isdir(self.path.get()):
				self.master.destroy()
			else:
				self.popup()

		def popup(self):
			self.w = popupWindow(self.master)
			self.b["state"] = "disabled" 
			self.master.wait_window(self.w.top)
			self.b["state"] = "normal"

		def browse_button(self):
			# current_focus = Toplevel()
			# current_focus.wm_attributes('-topmost', 1)
			# current_focus.withdraw()
			# current_focus.protocol('WM_DELETE_WINDOW', current_focus.withdraw)
			# oldFoc = current_focus.focus_get()
			path_ = filedialog.askdirectory()
			# if oldFoc: oldFoc.focus_set()
			self.path.set(path_)
			# self.path.set(path_.replace("/", "\\\\"))

	################## install_python358 #######################
	def install_python358():
		try:
			subprocess.call("\"" + current_installing_path.replace("\\", "/") + "/" + "python-3.6.8.exe" + "\"")
		except:
			print("install_python358 fail")
	################## install_python358 #######################

	python_excute_code = ""
	################## upgrade_pip_and_install_jupyter #######################
	def upgrade_pip_and_install_jupyter():
		global python_excute_code
		python_excute_code = ""
		try:
			python_excute_code = "python -m "
			print(python_excute_code + "pip install " + "--upgrade pip")
			subprocess.call(python_excute_code + "pip install " + "--upgrade pip")
		except:
			try:
				python_excute_code = "py -m "
				print(python_excute_code + "pip install " + "--upgrade pip")
				subprocess.call(python_excute_code + "pip install " + "--upgrade pip")
			except:
				print('"pip install --upgrade pip" fail')
				os.system("pause")
				exit()

		try:
			print(python_excute_code + "pip install " + "jupyter")
			subprocess.call(python_excute_code + "pip install " + "jupyter")
		except:
			print('"install jupyter" fail')
			os.system("pause")
			exit()
	################## upgrade_pip_and_install_jupyter #######################

	################## packages_installing #######################
	def packages_installing(value):
		global python_excute_code
		packages = value.split('\n')
		for t in packages[::-1]:
			if len(t) <= 2:
				packages.pop()
		packages_installing_Exception = []
		tensorflow_installing_success = False
		for package_name in packages:
			if package_name != "tensorflow":
				try:
					print(python_excute_code + "pip install " + package_name)
					subprocess.call(python_excute_code + "pip install " + package_name)
				except Exception as ex:
					packages_installing_Exception.append(ex)
			else:
				try:
					print(python_excute_code + "pip install " + package_name)
					subprocess.call(python_excute_code + "pip install " + package_name)
					tensorflow_installing_success = True
				except:
					pass
				try:
					print(python_excute_code + "pip3 install " + package_name)
					subprocess.call(python_excute_code + "pip3 install " + package_name)
					tensorflow_installing_success = True
				except:
					pass
				try:
					print(python_excute_code + "pip install " + "--upgrade https://storage.googleapis.com/tensorflow/windows/cpu/tensorflow-0.12.0rc0-cp35-cp35m-win_amd64.whl")
					subprocess.call(python_excute_code + "pip install " + "--upgrade https://storage.googleapis.com/tensorflow/windows/cpu/tensorflow-0.12.0rc0-cp35-cp35m-win_amd64.whl")
					tensorflow_installing_success = True
				except Exception as ex:
					tensorflow_installing_Exception = ex
					pass
				try:
					print(python_excute_code + "pip install " + "--upgrade -I setuptools")
					subprocess.call(python_excute_code + "pip install " + "--upgrade -I setuptools")
					print(python_excute_code + "pip install " + "--ignore-installed --upgrade https://storage.googleapis.com/tensorflow/windows/cpu/tensorflow-1.0.1-cp35-cp35m-win_amd64.whl")
					subprocess.call(python_excute_code + "pip install " + "--ignore-installed --upgrade https://storage.googleapis.com/tensorflow/windows/cpu/tensorflow-1.0.1-cp35-cp35m-win_amd64.whl")
					tensorflow_installing_success = True
				except:
					pass
				# pip install tensorflow
				# pip3 install tensorflow
				# pip install --upgrade https://storage.googleapis.com/tensorflow/windows/cpu/tensorflow-0.12.0rc0-cp35-cp35m-win_amd64.whl
				# pip install --upgrade -I setuptools
				# pip install --ignore-installed --upgrade https://storage.googleapis.com/tensorflow/windows/cpu/tensorflow-1.0.1-cp35-cp35m-win_amd64.whl
		for package_installing_Exception in packages_installing_Exception:
			print("\n")
			print(package_installing_Exception)
		if (not tensorflow_installing_success) and ("tensorflow" in packages):
			print("\n")
			print(tensorflow_installing_Exception)
		if (len(packages_installing_Exception) > 0) or ((not tensorflow_installing_success) and ("tensorflow" in packages)):
			os.system("pause")
	################## packages_installing #######################

	################## jupyter_notebook_config #######################
	def get_default_windows_app(suffix): # get_default_windows_app(".html")
		class_root = winreg.QueryValue(winreg.HKEY_CLASSES_ROOT, suffix)
		with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r'{}\shell\open\command'.format(class_root)) as key:
			command = winreg.QueryValueEx(key, '')[0]
			return shlex.split(command)[0]

	def create_and_setting_jupyter_notebook_config(jupyter_default_folder_path):
		global python_excute_code
		try:
			print("jupyter notebook --generate-config")
			subprocess.call("jupyter notebook --generate-config")
		except:
			try:
				print("notebook --generate-config")
				subprocess.call("notebook --generate-config")
			except:
				try:
					print(python_excute_code + "jupyter notebook --generate-config")
					subprocess.call(python_excute_code + "jupyter notebook --generate-config")
				except:
					try:
						print(python_excute_code + "notebook --generate-config")
						subprocess.call(python_excute_code + "notebook --generate-config")
					except:
						print("jupyter notebook --generate-config fail")
		Userprofile_path = os.environ['USERPROFILE']
		jupyter_notebook_config_path = os.path.join(Userprofile_path, ".jupyter", "jupyter_notebook_config.py")
		try:
			f = open(jupyter_notebook_config_path, "r")
		except:
			print("Opening the jupyter_notebook_config.py fail!")
			exit()
		jupyter_notebook_config = f.read()
		f.close()
		string_for_setting_dir = "c.NotebookApp.notebook_dir = r'"
		string_for_setting_dir += jupyter_default_folder_path
		string_for_setting_dir += "'"
		jupyter_notebook_config = jupyter_notebook_config.replace("# c.NotebookApp.notebook_dir = ''", string_for_setting_dir)
		string_for_setting_browser = "import webbrowser\n"
		string_for_setting_browser += "webbrowser.register('chrome', None, webbrowser.GenericBrowser('"
		string_for_setting_browser += get_default_windows_app(".html")
		string_for_setting_browser += "'))\n"
		string_for_setting_browser += "c.NotebookApp.browser = 'chrome'\n"
		jupyter_notebook_config = jupyter_notebook_config.replace("# c.NotebookApp.browser = ''", string_for_setting_browser)
		f = open(jupyter_notebook_config_path, "w")
		f.write(jupyter_notebook_config)
		f.close()

		Desktop_path = os.path.join(Userprofile_path, 'Desktop') # path to where you want to put the .lnk
		path = os.path.join(Desktop_path, 'jupyter_notebook.lnk')
		target = os.path.join(Userprofile_path, "AppData", "Local", "Programs", "Python", "Python36-32", "Scripts", "jupyter-notebook.exe")
		# icon = r'C:\path\to\icon\resource.ico' # not needed, but nice

		shell = win32com.client.Dispatch("WScript.Shell")
		shortcut = shell.CreateShortCut(path)
		shortcut.Targetpath = target
		# shortcut.IconLocation = icon
		shortcut.WindowStyle = 7 # 7 - Minimized, 3 - Maximized, 1 - Normal
		shortcut.save()

		# c.NotebookApp.notebook_dir = ''
		# c.NotebookApp.notebook_dir = r''

		# c.NotebookApp.browser = ''
		# c.NotebookApp.browser = 'c:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s'
		# import webbrowser
		# webbrowser.register('chrome', None, webbrowser.GenericBrowser('C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'))
		# c.NotebookApp.browser = 'chrome'
		# \Local\Programs\Python\Python36-32\Scripts
	################## jupyter_notebook_config #######################

	class endWindow(object):
		def __init__(self, master):
			self.master = master
			self.master.title("Python with jupyter notebook installing")
			self.master.geometry("100x100") #the size of the app
			self.master.resizable(0, 0) #Don't allow resizing in the x or y direction
			backGroundCanvas = Canvas(master, width = 100, height = 100)
			backGroundCanvas.focus_set()
			backGroundCanvas.pack()
			self.l = Label(self.master, text = "完成")
			self.l.place(x = 35, y = 20)
			# self.l.pack()
			self.b = Button(self.master, text = 'Ok', command = self.destroy_popup, height = 2, width = 8)
			self.b.place(x = 16, y = 40)
			# self.b.pack()

		def destroy_popup(self):
			self.master.destroy()


	if __name__ == "__main__":
		global packages_list, jupyter_default_folder_path
		install_python358()
		install_python358()
		root = Tk()
		m = mainWindow(root)
		root.mainloop()
		packages_list, jupyter_default_folder_path = m.entryValue()
		# print(packages_list)
		# print(jupyter_default_folder_path)
		if not os.path.isdir(jupyter_default_folder_path):
			print(jupyter_default_folder_path + "路徑不存在")
			os.system("pause")
			exit()
		upgrade_pip_and_install_jupyter()
		# print(python_excute_code)
		packages_installing(packages_list)
		create_and_setting_jupyter_notebook_config(jupyter_default_folder_path)
		end_root = Tk()
		end_m = endWindow(end_root)
		end_root.mainloop()

# finally:
	# os.system("pause")
except Exception as ex:
	print(ex)
	os.system("pause")
	# raw_input()


# jupyter-notebook.exe
################## create_shortcut #######################
# import win32com.client
# import os
# # import pythoncom # remove the '#' at the beginning of the line if running in a thread.
# # pythoncom.CoInitialize() # remove the '#' at the beginning of the line if running in a thread.
# Userprofile_path = os.environ['USERPROFILE']
# Desktop_path = os.path.join(Userprofile_path, 'Desktop') # path to where you want to put the .lnk
# path = os.path.join(Desktop_path, 'jupyter_notebook.lnk')
# target = os.path.join(Userprofile_path, "Local", "Programs", "Python", "Python36-32", "Scripts", "jupyter-notebook.exe")
# # icon = r'C:\path\to\icon\resource.ico' # not needed, but nice

# shell = win32com.client.Dispatch("WScript.Shell")
# shortcut = shell.CreateShortCut(path)
# shortcut.Targetpath = target
# # shortcut.IconLocation = icon
# shortcut.WindowStyle = 7 # 7 - Minimized, 3 - Maximized, 1 - Normal
# shortcut.save()
################## create_shortcut #######################
