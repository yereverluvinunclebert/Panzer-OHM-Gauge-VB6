# Panzer-OHM-Gauge-VB6
 
A FOSS Temperature Gauge VB6 WoW64 Widget for Windows Vista, 7, 8 and 10/11+. The Panzer OHM Gauge VB6 is a useful utility displaying the temperatures of all cores and components in a dieselpunk fashion on your desktop. There will also be a version later for Reactos and XP, watch this space for the link. Also tested and running well on Linux and Mac os/X using Wine.

![security](https://github.com/yereverluvinunclebert/Panzer-OHM-Gauge-VB6/assets/2788342/3fc22b45-ca08-4c2b-98f1-8cce6e4003c6)

One of my VB6/RC6 PSD WoW 64 designs, a work largely complete but as with all software some tasks to complete and unfound bugs to repair. However, it works fully, there is an installer and it does the job. New version always coming.  This VB6 widget is based upon the Yahoo/Konfabulator and Xwidgets of the same design. Those older widget platforms are now obsolete so I have migrated the code to a new development environment. Each exists now as a Windows binary rather than a widget requiring a runtime. 

Download it here: https://github.com/yereverluvinunclebert/Panzer-OHM-Gauge-VB6/releases/tag/0.1.192beta

 Why VB6? Well, with a 64 bit, modern-language improvement upgrade on the way with 100% compatible TwinBasic coupled with support for transparent PNGs via RC/Cairo, VB6 code has an amazing future.

![vb6-logo-350](https://github.com/yereverluvinunclebert/Panzer-CPU-Gauge-VB6/assets/2788342/39e2c93f-40a5-4c47-8c23-d8ce7c747b10)

 I created as a variation of the previous gauges I had previously created for the World of Tanks and War Thunder 
 communities. The Panzer OHM Gauge widget is an attractive dieselpunk VB6 widget for your desktop. 
 Functional and gorgeous at the same time. The graphics are my own, I took original inspiration from a clock face by Italo Fortana combining it with an aircraft gauge. It is all my code with some help from the chaps at VBForums (credits given). 
  
This Widget is a moveable widget that you can move anywhere around the desktop as you require. Requires installation of Open Hardware Monitor or Libre Hardware Monitor to allow the temperatures to be extracted and written to the WMI data repository. The Panzer OHM Gauge VB6 extracts the temperature data from that repository and displays it as a Panzer gauge on the desktop. 

This is the pertinent bit:

    Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objSWbemServices = objSWbemLocator.ConnectServer(strComputer, "root\OpenHardwareMonitor")
    Set colItems = objSWbemServices.ExecQuery("SELECT * FROM Sensor WHERE SensorType = 'Temperature'")

In time there will also be coreTemp upgrade soon - watch this space!
 
![panzer-OHMphoto-1440x900](https://github.com/yereverluvinunclebert/Panzer-OHM-Gauge-VB6/assets/2788342/f44dea4e-ed05-44a4-8516-3c06eaa01068)

 This widget can be increased in size, animation speed can be changed, 
 opacity/transparency may be set as to the users discretion. The widget can 
 also be made to hide for a pre-determined period.
 
![panzerOHMgauge](https://github.com/yereverluvinunclebert/Panzer-OHM-Gauge-VB6/assets/2788342/a2172003-befd-4ec3-a9f2-9f0ef69416d7)

 Right clicking will bring up a menu of options. Double-clicking on the widget will cause a personalised Windows application to 
 fire up. The first time you run it there will be no assigned function and so it 
 will state as such and then pop up the preferences so that you can enter the 
 command of your choice. The widget takes command line-style commands for 
 windows. Mouse hover over the widget and press CTRL+mousewheel up/down to resize. It works well on Windows Vista/7 to Windows 11+.
 
![panzer-temperature-OHM-help](https://github.com/yereverluvinunclebert/Panzer-OHM-Gauge-VB6/assets/2788342/55c5a56f-2ba6-4dad-8649-d8a9314cc9d5)
 
 The Panzer OHM Gauge VB6 gauge is Beta-grade software, under development, not yet 
 ready to use on a production system - use at your own risk.

 This version was developed on Windows 7 using 32 bit VisualBasic 6 as a FOSS 
 project creating a WoW64 widget for the desktop. 
 
![Licence002](https://github.com/yereverluvinunclebert/Panzer-CPU-Gauge-VB6/assets/2788342/a24c5c45-5517-4423-938b-398f1d349d4c)

 It is open source to allow easy configuration, bug-fixing, enhancement and 
 community contribution towards free-and-useful VB6 utilities that can be created
 by anyone. The first step was the creation of this template program to form the 
 basis for the conversion of other desktop utilities or widgets. A future step 
 is new VB6 widgets with more functionality and then hopefully, conversion of 
 each to RADBasic/TwinBasic for future-proofing and 64bit-ness. 

![menu01](https://github.com/yereverluvinunclebert/Panzer-CPU-Gauge-VB6/assets/2788342/45da67ae-8ec9-4730-a00d-e3c5a1db7519)

 This utility is one of a set of steampunk and dieselpunk widgets. That you can 
 find here on Deviantart: https://www.deviantart.com/yereverluvinuncleber/gallery
 
 I do hope you enjoy using this utility and others. Your own software 
 enhancements and contributions will be gratefully received if you choose to 
 contribute.

![security](https://github.com/yereverluvinunclebert/Panzer-OHM-Gauge-VB6/assets/2788342/3fc22b45-ca08-4c2b-98f1-8cce6e4003c6)

 BUILD: The program runs without any Microsoft plugins.
 
 Built using: VB6, MZ-TOOLS 3.0, VBAdvance, CodeHelp Core IDE Extender
 Framework 2.2 & Rubberduck 2.4.1, RichClient 6
 
 Links:
 
	https://www.vbrichclient.com/#/en/About/
	MZ-TOOLS https://www.mztools.com/  
	CodeHelp http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=62468&lngWId=1  
	Rubberduck http://rubberduckvba.com/  
	Rocketdock https://punklabs.com/  
	Registry code ALLAPI.COM  
	La Volpe http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=67466&lngWId=1  
	PrivateExtractIcons code http://www.activevb.de/rubriken/  
	Persistent debug code http://www.vbforums.com/member.php?234143-Elroy  
	Open File common dialog code without dependent OCX - http://forums.codeguru.com/member.php?92278-rxbagain  
	VBAdvance  
 
 
 Tested on :
 
	ReactOS 0.4.14 32bit on virtualBox    
	Windows 7 Professional 32bit on Intel    
	Windows 7 Ultimate 64bit on Intel    
	Windows 7 Professional 64bit on Intel    
	Windows XP SP3 32bit on Intel    
	Windows 10 Home 64bit on Intel    
	Windows 10 Home 64bit on AMD    
	Windows 11 64bit on Intel  
   
 CREDITS:
 
 I have really tried to maintain the credits as the project has progressed. If I 
 have made a mistake and left someone out then do forgive me. I will make amends 
 if anyone points out my mistake in leaving someone out.
 
 MicroSoft in the 90s - MS built good, lean and useful tools in the late 90s and 
 early 2000s. Thanks for VB6.
 
 Olaf Schmidt - This tool was built using the RichClient RC5 Cairo wrapper for 
 VB6. Specifically the components using transparency and reading images directly 
 from PSD. Thanks for the massive effort Olaf in creating Cairo counterparts for 
 all VB6 native controls and giving us access to advanced features on controls 
 such as transparency.
 
 Shuja Ali @ codeguru for his settings.ini code.
 
 ALLAPI.COM        For the registry reading code.
 
 Rxbagain on codeguru for his Open File common dialog code without a dependent 
 OCX - http://forums.codeguru.com/member.php?92278-rxbagain
 
 si_the_geek       for his special folder code
 
 Elroy on VB forums for the balloon tooltips
 
 Harry Whitfield for his quality testing, brain stimulation and being an 
 unwitting source of inspiration. 
 
 Dependencies:
 
 o A windows-alike o/s such as Windows XP, 7-11 or Apple Mac OSX 11. 
 
 o Microsoft VB6 IDE installed with its runtime components. The program runs 
 without any additional Microsoft OCX components, just the basic controls that 
 ship with VB6.  
 
 ![vb6-logo-200](https://github.com/yereverluvinunclebert/Panzer-CPU-Gauge-VB6/assets/2788342/bf00fa3d-f1d4-417b-bc50-9446f2c3e674)

 
 * Uses the latest version of the RC6 Cairo framework from Olaf Schmidt.
 
 During development the RC6 components need to be registered. These scripts are 
 used to register. Run each by double-clicking on them. They are in the BIN folder.
 
	RegisterRC6inPlace.vbs
	RegisterRC6WidgetsInPlace.vbs
 
 During runtime on the users system, the RC6 components are dynamically 
 referenced using modRC6regfree.bas which is compiled into the binary.	
 
 
 Requires a PzOHM Gauge folder in C:\Users\<user>\AppData\Roaming\ 
 eg: C:\Users\<user>\AppData\Roaming\PzOHM Gauge
 Requires a settings.ini file to exist in C:\Users\<user>\AppData\Roaming\PzOHM Gauge
 The above will be created automatically by the compiled program when run for the 
 first time.
 
o Krool's replacement for the Microsoft Windows Common Controls found in
mscomctl.ocx (slider) are replicated by the addition of one
dedicated OCX file that are shipped with this package.

During development only, this must be copied to C:\windows\syswow64 and should be registered.

- CCRSlider.ocx

Register this using regsvr32, ie. in a CMD window with administrator privileges.
	
	c:                          ! set device to boot drive with Windows
	cd \windows\syswow64s	    ! change default folder to syswow64
	regsvr32 CCRSlider.ocx	! register the ocx

This will allow the custom controls to be accessible to the VB6 IDE
at design time and the sliders will function as intended (if this ocx is
not registered correctly then the relevant controls will be replaced by picture boxes).

![ccrslider](https://github.com/user-attachments/assets/b22e6b3e-e351-4d24-be90-ce44305b5f4f)

If you are developing multiple widgets needing this ocx, the above only needs to be done once. 

The above rigmarole is only for development. For ordinary users during runtime, there is no need to do the above. The OCX will reside in the program folder. The program reference to this OCX is contained within the supplied resource file, Panzer OHM Gauge.RES. The reference to this file is already compiled into the binary. As long as the OCX is in the same folder as the binary the program will run without the need to register the OCX manually.
 
 
![ohmslider](https://github.com/user-attachments/assets/00fc914e-ea69-4649-8579-42e5a0185f57)

Sample preference screen showing CCR slider in place and operating

 * SETUP.EXE - The program is currently distributed using setup2go, a very useful 
 and comprehensive installer program that builds a .exe installer. Youll have to 
 find a copy of setup2go on the web as it is now abandonware. Contact me
 directly for a copy. The file "install PzOHM Gauge 0.1.0.s2g" is the configuration 
 file for setup2go. When you build it will report any errors in the build.
 
 * HELP.CHM - the program documentation is built using the NVU HTML editor and 
 compiled using the Microsoft supplied CHM builder tools (HTMLHelp Workshop) and 
 the HTM2CHM tool from Yaroslav Kirillov. Both are abandonware but still do
 the job admirably. The HTML files exist alongside the compiled CHM file in the 
 HELP folder.
 
  Project References:

	VisualBasic for Applications  
	VisualBasic Runtime Objects and Procedures  
	VisualBasic Objects and Procedures  
	OLE Automation  
 	Microsoft WMI Scripting V1.2 Library wbemdisp.tlb 
	vbRichClient6  	- RC6Widgets (RC6Widgets.DLL)
                   	- RC6 (RC6.DLL)

 ![wmi](https://github.com/user-attachments/assets/bfc329ad-4ca4-4552-a05a-6f3ccc003269)

 
 LICENCE AGREEMENTS:
 
 Copyright © 2023 Dean Beedell
 
 In addition to the GNU General Public Licence please be aware that you may use 
 any of my own imagery in your own creations but commercially only with my 
 permission. In all other non-commercial cases I require a credit to the 
 original artist using my name or one of my pseudonyms and a link to my site. 
 With regard to the commercial use of incorporated images, permission and a 
 licence would need to be obtained from the original owner and creator, ie. me.
 
![about](https://github.com/yereverluvinunclebert/Panzer-OHM-Gauge-VB6/assets/2788342/e48a2165-efa1-44ae-a75a-46801ea461ab)

![Panzer-CPU-Gauge-onDesktop](https://github.com/yereverluvinunclebert/Panzer-CPU-Gauge-VB6/assets/2788342/f2e06e77-77ee-46fb-840d-e14f96b4e0a5)





 
