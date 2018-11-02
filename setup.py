from cx_Freeze import setup, Executable

setup(name="ParseInbox",
      version="0.1",
      description="",
      executables=[Executable("AppParseOutlookFolders.py")], requires=['gevent'])