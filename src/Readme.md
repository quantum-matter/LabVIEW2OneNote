## Python scripts for LabVIEW2OneNoteController

This set of Python scripts serve as the underlying code that is called by the LabVIEW VI wrappers when performing programmatic control of OneNote to add text/images/plots from inside LabVIEW.

### Instructions for using the Python scripts:
* Update the set of credentials near the top of the source code in each script, after obtaining said credentials from Microsoft Azure, as has been described in the [Documentation PDF.](https://github.com/quantum-matter/LabVIEW2OneNote/blob/master/LabVIEW%20for%20OneNote%202016%20-%20Documentation.pdf "Documentation - LabVIEW2OneNoteController")
* Compile the `.py` files into `.exe` using [`pyinstaller`](https://www.pyinstaller.org/ "PyInstaller") and place them along with your LabVIEW library. The VIs are designed to work with executables and not scripts.
* Once the credentials have been updated, the Python scripts can also be used independently without compiling or using LabVIEW VI wrappers in the following manner:
  - `python createTodayPage.py` *Notebook_Name*
  - `python addHeaderTodayPage.py` *Notebook_Name* *Font_Size* 
  - `python addImageTodayPage.py` *Notebook_Name* (assuming that local server is running and tunneled through [`ngrok`](https://ngrok.com/ "ngrok"), and image named `tempImg.png` is present in server root directory)
  - `python addTextTodayPage.py` *Notebook_Name* *Font_Size* *Text_in_Quotes*
  - `python addTableTodayPage.py` *Notebook_Name* (assuming `tempTable.csv` and `widthFile.csv` are present in the same directory)
* The `addImageTodayPage.py` script (and any LabVIEW wrappers calling it) requires that the `ngrok.exe` file (can be downloaded from [here](https://ngrok.com/ "ngrok")) be present in the same directory as the `.llb` and the `.exe` files. The VIs are designed to configure it with a pre-supplied `ngrok` Auth Token, but it is advised to register for your own free Auth Token.
* All these Python scripts will automatically first create a New Page for the Day in the Chronologic structure in the provided OneNote Notebook if it doesn't already exist
* The `Python_server.py` is a simple `http` server implementation that makes use of Python's built-in `SimpleHTTPRequestHandler`. To start a server in the same directory, use `python Python_server.py` *port_no*. If no `port_no` is supplied, it starts at `7800`.
