from os import path, environ, listdir
from ctypes import windll
from random import choice
from sys import argv
from win32com.client import Dispatch

def createStarupShortcut(file_path):
    dir_path = path.dirname(file_path)
    file_name = path.basename(file_path)
    file_name_without_extension = path.splitext(file_name)[0]

    startup_path = environ['HOMEPATH'] + f'\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup'
    shortcut_path = path.join(startup_path, f"{file_name_without_extension}.lnk")

    if not(path.isfile(shortcut_path)):
        shell = Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = file_path
        shortcut.WorkingDirectory = dir_path
        shortcut.IconLocation = file_path
        shortcut.save()

def setWallpaper(path):
    windll.user32.SystemParametersInfoW(20, 0, path, 0)

def main():
    if(path.isdir('wallpapers')):
        file_path = argv[0]

        wallpapers_path = path.join(path.dirname(path.abspath(file_path)), 'wallpapers')
        wallpapers = listdir(wallpapers_path)

        if(len(wallpapers) > 0):
            random_wallpaper_path = path.join(wallpapers_path, choice(wallpapers))
            
            setWallpaper(random_wallpaper_path)

            createStarupShortcut(file_path)

if __name__ == "__main__":
    main()   
            