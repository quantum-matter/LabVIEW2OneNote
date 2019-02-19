## Python scripts for LabVIEW2OneNoteController

This set of Python scripts serve as the underlying code that is called by the LabVIEW VI wrappers when performing programmatic control of OneNote to add text/images/plots from inside LabVIEW.

### Instructions on using the Python scripts:
* Update the set of credentials near the top of the source code in each script, after obtaining said credentials from Microsoft Azure, as has been described in the Documentation
* Compile the `.py` files into `.exe` using [`pyinstaller`](https://www.pyinstaller.org/ "PyInstaller") and place them along with your LabVIEW library. The VIs are designed to work with executables and not scripts.
