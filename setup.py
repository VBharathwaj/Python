from cx_Freeze import setup, Executable

setup(name = "CheckappReport" ,
      version = "0.1" ,
      description = "" ,
      executables = [Executable("azure_upload_status.py")])