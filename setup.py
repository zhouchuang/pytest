from distutils.core import setup
import  py2exe
import  PyInstaller

options={"py2exe":{"bundle_files":1}
        }
setup(options=options,
      zipfile=None,
      console=['test63.py'])
