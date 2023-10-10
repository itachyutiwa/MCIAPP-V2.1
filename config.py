from cx_Freeze import setup, Executable

setup(
    name="MCI-CARE-CI",
    version="1.0",
    description="Application de Contrôle Délégué MCI-CARE-CI, Member Of Sanlam Assurance CI",
    executables=[Executable("app.py", base="Win32GUI")],
    options={
        'build_exe': {
            'packages': ['tkinter','ttkbootstrap','pandas','numpy','openpyxl'],
            'include_files': [("./images/mci_care.png", './images/mci_care.png')],
        },
    },
)
