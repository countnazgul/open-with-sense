### Work In Progress

"Open with Sense" is a small console application that will adds context menu option to quickly load `csv` and `qvd` files in Qlik Sense Desktop. 

### Limitations in the initial version
  * works only with Qlik Sense Desktop
  * works only with QS version 3.2+
  * `legacy mode` should be enabled
  * only `csv` and `qvd` files can be loaded
  
### How-to

#### Install / Uninstall
 There is no installation process (for now). Just download the release zip file and extract it somewhere. Start the "OpenWithSense.exe" (make sure QS Desktop is running) and pick option `2`. This will make the necesarry changes in the Windows Registry and to the system `PATH` variable.
 
 To uninstall it - start "OpenWithSense" from within command prompth and choose option `3` then delete the folder where the app is.

#### Usage
  * Right click on any file (if the file is not csv, qvd or qvf the app will exit) and choose "Load in Qlik Sense Desktop". This will create new app, edit the script, reload, create sheet, save and open the app in the browser on the sheet in edit mode
  * The app itself provide few options
    * 1 - delete all documents created with the app
    * 2 - add right click options to registry
    * 3 - remove right click options from registry
    * 4 - open this page
  * QVF files - copy the targeted qvf file to QS Apps folder and opens it in the browser

### Future development
  * support QS Server (Qlik Cloud?)
  * no legacy mode - analyze and create new data connectors if needed (also to provide option to delete the "temp" connectors)
  * support more file types - excel, xml
  * builds for different (older) versions of QS
  * more options in the context menu (basically duplicate the options from the console app)
  * app templates - define your own design templates (qvf). For example: if the loaded data have geo information the template might contains separate sheet with already prepared geo related charts
