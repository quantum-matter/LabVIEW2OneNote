# LabVIEW2OneNote
#### LabVIEW Library for OneNote 2016

### Programmatic control of OneNote from LabVIEW with automated client authentication

This set of VIs and associated Python scripts allow you to to send information like images and text to a page in Microsoft OneNote 2016 through a LabVIEW interface to ensure seamless data journaling during experiments.

### Getting Started
To get started, clone the repository using `git clone` or dowmload it as a `zip` archive. The LabVIEW Library `LabVIEW2OneNoteController.llb` contains all the required VIs while the [`src`](https://github.com/quantum-matter/LabVIEW2OneNote/tree/master/src "LabVIEW2OneNoteController src") has the necessary Python scripts. Then obtain credentials, update the Python scripts and compile them into executables using [`pyinstaller`](https://www.pyinstaller.org/ "PyInstaller"). Some basic context is provided below and more instructions are available in the [Documentation PDF](https://github.com/quantum-matter/LabVIEW2OneNote/blob/master/LabVIEW%20for%20OneNote%202016%20-%20Documentation.pdf "Documentation - LabVIEW2OneNoteController").

Copy the executables into the folder containing the `llb` files. Then run the `LabVIEW2OneNote example.llb` library and make sure that all the VIs get correctly loaded from the `LabVIEW2OneNoteController.llb`. You can then send text/tables/images to OneNote from LabVIEW.

### Background
In our case the purpose is to let LabVIEW assist us in writing a chronologic OneNote journal (with year, month, date structure) for the data generated from our LabVIEW experimental control system. In former solutions for this task we used the ActiveX library of OneNote, and later the DDE, both for Microsoft Office 2013 and before. Now we have implemented this for OneNote 2016 & above, bundled with Microsoft Office 365. This makes use of the `REST`-based [Microsoft Graph API](https://developer.microsoft.com/en-us/graph "Microsoft Graph API"), which allows integration across the suite of Microsoft apps with similar API endpoints. Given that our experimental control system in LabVIEW is quite a large application, we have not recently updated it and are still running under LabVIEW 2010 (full development system). The VIs and Python scripts in this repository were designed keeping in mind our very specific use-case of maintaining a Laboratory journal while performing experimental control through LabVIEW. As such, various features/data-flows might come across as quirks but they often latch onto and integrate with legacy code for niche usage. However, both the Python scripts and the LabVIEW VIs are extremely modular and contain sufficient documentation to be extended to any possible server-based controlling of your OneNote notebooks.

Programmatic access to OneNote for Office 365 is based on a Microsoft Office 365 account that is used to grant the required permissions to the app via the [Azure Active Directory](https://docs.microsoft.com/en-us/azure/active-directory/fundamentals/active-directory-whatis "What is Azure Active Directory?") feature on [Microsoft Azure Portal](https://portal.azure.com "Microsoft Azure Portal"), from which a set of `OAuth 2.0` access credentials is generated. The various endpoints of the Graph OneNote API are then exposed as further described [here](https://docs.microsoft.com/en-us/graph/integrate-with-onenote "OneNote API overview") and [here](https://docs.microsoft.com/en-gb/graph/auth-v2-service "Get access without a user"). The `http` REST calls are then made from Python modules which are compiled into executables for portability and machine independence. The compiled Python scripts are then called by corresponding wrapper VIs that make the functionality available in LabVIEW. 

### Environment
`Python 3.6+`
`LabVIEW 2010+`
`Microsoft OneNote 2016+ with Office 365`
