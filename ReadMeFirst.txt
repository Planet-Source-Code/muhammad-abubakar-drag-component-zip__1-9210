~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Hi,
Drag Component <DragX.dll> is my first ActiveX DLL. It can enable any application to accept files that are dragged over it. Its very easy to use in your applications.

METHODS AND PROPERTIES:

DragHwnd
--------
->Property. Handle of the window over which you want to capture the draged files. It can be any window like ListBox, TreeView, TextBox etc.

FileCount
---------
->Property.Read Only. Contains the number of files droped.

FileName
--------
->Method. Gets the filenames. Suppose 5 files were droped over a window then 
	object.FileName(3)
will get the full path filename of the fourth file. It works like an index which is 0 (zero) based (starts with zero). The maximum value that can be given to FileName is : FileCount - 1 .

StartDrag
---------
->Method. Starts the subclassing and monitores the window (handle of which is given to DragHwnd property) for WM_DROPFILES.

StopDrag
--------
->Method. Stops tracking for the files dropped over a window.

FilesDroped
-----------
->Event. This event is raised as soon as the file(s) are dropped (WM_DROPFILES message encountered).

CONTROLLING SUBCLASSING FROM WITHIN A CLASS:
---------------------------------------------
The process of enabling your applications to accept draging files from Windows Explorer is to subclass the window, over which you expect the files to be draged, intercepting the WM_DROPFILES message and using several APIs to query for the files droped. When you call the 'DragStart' method it changes the default procedure of the window, whose handle is given to 'DragHwnd' property , by using 'SetWindowLong' and from that moment on all messages concerning subclassed window are passed to the 'WndProc' function whose address in given to 'SetWindowLong'. If files are draged and droped over the hooked window, 'WndProc' recieves a WM_DROPFILES message. Now that 'WndProc' is located in a BAS module (we have no choice to place it anywhere else since its the condition of 'AddressOf' operator that the subclassed procedure must reside in a BAS file) we need a way to communicate the names of the dropped files to the class. But be careful, the 'functions' that we'll be using to communicate between a class and a module are not to be shown to the users using the class. We want the module to see some of the functions of the class, so you might think of making those functions 'Public', but no, that'll make them class's methods which'll be of no use to users of the class and will confuse them because they'll be able to see it and do nothing with it. So we have a need of those special kind of procedures that are public but are not class members at the same time, hence a need for 'Friend' functions...

FRIENDLY 'FRIEND' FUNCTIONS :^)
----------------------------
Friend functions are not considered class members. When a method is declared with a Friend keyword, its visible to other objects in your component, but its not called using member selection operator (.) and its not added to the type library or the public interface. I'm using three friendly functions in my class:

i)	ClearFileNames
ii)	AddInFileNames
iii)	NowRaiseEvent

When 'WndProc' gets a WM_DROPFILES message it calls the 'ClearFileNames' to clear the previously dropped files data and then gets the name of each file (with path) and passes it to the 'AddInFileNames' which stores the name to FileNames() array. When all names are passed, the 'WndProc' calls third friendly function 'NowRaiseEvent' which in turns raises 'FilesDroped' event.

PROJECT FILES:
--------------
The zip file that you've downloaded contains a project group 'Drag_Component.vbg', the DragX.dll , and Project1.vbp. 'Drag_Component' is the source code of the DLL and contains a TestProject to test the DLL. The 'Class Version' folder contains a project which shows the use of Drag component as a class in your project. You must register the DragX.dll before you can use it. An easy way to register the DragX.dll is to goto Project menu and selecting References and browse all your way to the place where you've extracted the DLL and selecting it, VB will register it. The class version project has a notepad sort of program called 'XFiles' which has the ability to open files by draging the files over it. Actually its the listbox above the RichTextBox over which you have to drag the file to open. Using the XFiles.exe you can open files in three ways, a) when XFiles is running and you drag the icon of the file (with extensions like *.sql,*.cpp,*.cls etc) over the listbox, b) the standard way i.e., by selecting Open from file menu , c) XFiles is not running but you just drag the icon of the targeting file over the icon of XFiles will automatically open with the specified file. The XFiles open and saves the text written on it in rtfText format, if you want it to work as rtfRTF format then you can change that in the 'FSave' and 'FLoad' subs in the form code window.

'Drag_Component' vs 'Drag Class':
---------------------------------
Now I'll like to argue a little about why should you prefer using the Drag component as a DLL rather then using it as a class in your project.
First. Using the Drag component as a DLL and including it in the setup files is basically following the philosophy of COM. Suppose you've shipped your app to the clients and later on a bug is reported and you find the bug was in Drag component. You'll debug it. Now if your app included a separate DLL then all you need is just send the clients an upgraded version of DLL and that's all, but if you have hard coded the class within your EXE then you'll have to ship the entire EXE again!
Second. Suppose you are using the CDrag_Drop class in your project (not the DLL) and its working, if your code has a bug somewhere and the compiler gets the bug and enters the [break] mode, VB will crash because VB IDE cannot work in the break mode while subclassing is going on. And just because of the Drag class your app debugging will become a hell for you (unless you are using 'Debug Object for AddressOf Subclassing'). However if you just had a reference to the Drag component DLL and you enter the [break] mode while debugging your app, everything will be just fine (and less amount of code windows around).

So now that you've an easy to use Drag Component you can make most of your apps Drag N Drop enabled like make your MDIs Drag N Drop enabled adding more functionality and making them more user friendly.

If you have any questions or comments regarding any of my programs please contact me at my email address <joehacker@yahoo.com> my website is http://go.to/abubakar. Thankx for downloading my program :^)
Regards.
 - Abubakar.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~