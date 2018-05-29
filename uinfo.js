try {
	///////////////////////SCRIPT SETTINGS//////////////////////////////
	var note = "" //leave a note (ENG characters only)
	var copydrive = "D" //logical drive can be changed automatically if is not enough space on selected drive and variable autochangecopydrive (below) is set to true
	var autochangecopydrive = false; //automatically changing if is not enough space on selected drive
	var needtocopy = false; //change to true if need to copy all finded files
	var autocopy = false; //start copying automatically or just create bat file for that
	var foldertocopyname = "findedfiles" //name of folder where files will been copied (plus _username)
	var needtoarchive = false; //change to true if need to create archive with all finded files
	var autoarchive = false; //start archiving automatically or just create bat file for that
	var archivename = "archive" //name of archive
	//var archiver = "\\\\10.111.110.10\\vol3\\@Obmen\\zip\\7za.exe" //archive programm path
	var archiver = "\\\\10.111.110.10\\vol3\\@Obmen\\zip\\7za.exe" //archive programm path
	var compressionlevel = 0 //for 0 to 9 (0-without compression 9-ultra compression)
	////////////////////////////////////////////////////////////
	var filecheckdelay = 1400; //Delay before next file and programs check (in min)//// 43200 (1 month)///
	var waittime = 0; // wait before start (in min)
	var config_1c_ext = "v8i"
var ext = [
"txt"
]
/*
	var ext = [ //files extensions for search
	"v8i", //1c configs
	"eml", //windows live mail
	"dbx", //outlook express
	"pst","ost", //ms outlook
	"tbb", //the bat
	"msf", //thunderbird
	"txt", //text files
	"rtf","doc","docx","docm","dot","dotm","dotx","odt", "wps", //ms word
	"xls","xlsx","csv","dbf","dif","ods","prn","slk","xla","xlam","xlsb","xlsm","xlt","xltm","xltx","xlw","xml","xps", //ms excel
	"ppt","pptx","pot","potm","potx","ppa","ppam","pps","ppsm","ppsx","pptm", //ms powerpoint
	"mdb","accdb", //ms access
	"rar","zip","arj","7z", //archives
	"ttf", //fonts
	"pdf", //adobe acrobat
	"djvu", //djvu
	"cda","wav","wma","mp3","avi","mpg","mpeg","mdv","flv","swf","divx","wmv", //media
	"bmp","gif","jpg","jpeg","tiff","png", //pictures
	"iso","mdf","mds","bin","nrg", //drive images
	"dwg","dfx","dgn","stl","dwt", //autocad
	"cdw","cdt","m3d","a3d", //compas 
	"vsd","vss","vst","vdx","vsx","vtx","vsl","vsdx","vsdm" //visio
	]; 
*/
	////////////////////////////////////////////////////////////
	////////////////////SCRIPT START////////////////////////////
	////////////////////////////////////////////////////////////

	WScript.Sleep ((waittime*1000)*60)
	var error = "";		
	var wmi = GetObject("winmgmts:\\\\.\\ROOT\\CIMV2");
	var wshell = WScript.CreateObject ("WScript.Shell");
	var fso = WScript.CreateObject ("Scripting.FileSystemObject");
	var timekey = "HKCU\\Software\\CheckTime";	
	var odinccheckkey = "HKCU\\Software\\Check1C";	
	var starttime = new Date();
	var file_checked = false;
	var odinconfigured = false;
	var file_finded = false;
	var programs_checked = false;
	///////////////////Time of last file checking/////////////////
	try {
		var value = null;
		try {value = wshell.RegRead(timekey)} catch(e) {}
		if(null != value){
			if ((((starttime.getTime() - (Number(wshell.RegRead(timekey))))/1000)/60) <= filecheckdelay) {
				file_checked = true;
			}
		}
	} catch(e) {error += "TIME|"}
	///////////////////Check is 1c configured on network configs/////////////////
	try {
		var value = null;
		try {value = wshell.RegRead(odinccheckkey)} catch(e) {}
		if(null != value){
			if (wshell.RegRead(odinccheckkey) == 1) {
				odinconfigured = true;
			}
		}
	} catch(e) {error += "TIME|"}
	///////////////////////////////NAME////////////////////////
	var user = "unknown";
	var compname = "unknown";
	var profile = "unknown";
	var username = "unknown";
	try {
		comp = (wshell.ExpandEnvironmentStrings("%COMPUTERNAME%"));
		username = (wshell.ExpandEnvironmentStrings("%USERNAME%"));
		domain = (wshell.ExpandEnvironmentStrings("%USERDOMAIN%"));
		profile = (wshell.ExpandEnvironmentStrings("%USERPROFILE%"));
		user = domain+"\\"+username
		compname = domain+"\\"+comp
		var usern = user.replace( /\\/g, "_user-" );
	} catch(e) {error += "NAME|"}
	
	
	///////////////////////////////OS////////////////////////
	var osStr = "unknown";
	var Arch = "unknown";
	var OS_Type = 0;
	try {
		query = wmi.ExecQuery("SELECT * FROM Win32_OperatingSystem");
		var os = new Enumerator(query).item();
		osStr = os.Caption + " " + os.Version
		Arch = os.OSArchitecture
		OS_Type=os.ProductType
	} catch(e) {error += "OS|"}
	///////////////////////////////RAM////////////////////////
	var RAM = 0;
	try {
		if (OS_Type == 1) {
		query = wmi.ExecQuery("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem");
		sys = new Enumerator(query).item();
		RAM = Math.round((sys.TotalPhysicalMemory/1024)/1024)
		}
	} catch(e) {error += "RAM|"}

	///////////////////////////////CPU//////////////////////////
	var CPU_name = "unknown";
	var CPU_freq = 0;
	try {
		if (OS_Type == 1) {
		query_cpu = wmi.ExecQuery("SELECT Name,MaxClockSpeed FROM Win32_Processor");
		cpu_arch = new Enumerator(query_cpu).item();
		  CPU_name = cpu_arch.Name
		  CPU_freq = cpu_arch.MaxClockSpeed
		}
	} catch(e) {error += "CPU|"}

	///////////////////////////////GPU//////////////////////////
	var GPU_NAME = "unknown";
	var GPU_RAM = 0;
	var GPU_HR = 0;
	var GPU_VR = 0;
	try {
		if (OS_Type == 1) {
			query_gpu = wmi.ExecQuery('SELECT Description,AdapterRAM,CurrentHorizontalResolution,CurrentVerticalResolution FROM Win32_VideoController WHERE NOT Description LIKE "%DameWare%"');
			gpu = new Enumerator(query_gpu);
			for(;!gpu.atEnd();gpu.moveNext()){
					if(((gpu.item().Availability)==3) || ((gpu.item().CurrentHorizontalResolution > 0) && (gpu.item().CurrentVerticalResolution > 0))){
						GPU_NAME = gpu.item().Description
						GPU_RAM = (Math.round((gpu.item().AdapterRAM/1024)/1024))
						GPU_HR = gpu.item().CurrentHorizontalResolution
						GPU_VR = gpu.item().CurrentVerticalResolution
					}else{
							GPU_NAME = gpu.item().Description
							GPU_RAM = (Math.round((gpu.item().AdapterRAM/1024)/1024))
							query_mon = wmi.ExecQuery("SELECT ScreenHeight,ScreenWidth FROM Win32_DesktopMonitor");
							monitor = new Enumerator(query_mon);
							for(;!monitor.atEnd();monitor.moveNext()){
								if ((monitor.item().Availability == 3) || ((monitor.item().ScreenHeight > 0) && (monitor.item().ScreenWidth > 0))) {
									GPU_HR = monitor.item().ScreenHeight
									GPU_VR = monitor.item().ScreenWidth
								}
							}
						 } 
			}	
		}
	} catch(e) {error += "GPU|"}
	
	///////////////////////////////Printers//////////////////////
	var printers_inf = ""
	try {
		if (OS_Type == 1) {
			query_printers = wmi.ExecQuery('SELECT Name,Default,Network,ServerName,ShareName FROM Win32_Printer WHERE (NOT Name LIKE "%Microsoft%" AND NOT Name LIKE "%Fax%" AND NOT Name LIKE "%XPS%" AND NOT Name LIKE "%Факс%" AND NOT Name LIKE "%PDF%" AND NOT Name LIKE "%OneNote%")');	
			printer = new Enumerator(query_printers);
			for(;!printer.atEnd();printer.moveNext()){
				printers_inf+="{"
				printers_inf+='"name":"' + printer.item().Name +'",'
				printers_inf+='"default":"'+ printer.item().Default +'",'
				printers_inf+='"network":"' + printer.item().Network +'",'
				printers_inf+='"server":"' + printer.item().ServerName +'",'
				printers_inf+='"sharename":"' + printer.item().ShareName +'"'
				printers_inf+="},"
			}
			printers_inf = printers_inf.slice(0,-1);
		}
	}catch(e) {error += "PRNT|"}
	
	//////////////////////////////Programs//////////////////////
	var programs_inf = ""
	var progs_array = []
	var arch_arr = []
	try {
		if ((OS_Type == 1) && (!file_checked)){
			/////////////////////get_from_WMI/////////////////
			query_programs = wmi.ExecQuery("SELECT Name FROM Win32_Product");	
			program = new Enumerator(query_programs);
			for(;!program.atEnd();program.moveNext()){
				var search_mask = program.item().Name.toLowerCase().replace(/\s/g, "").replace(/\+/g, "")
				if (((search_mask).search("\\(kb") == -1) && 
				((search_mask).search("visual") == -1) && 
				((search_mask).search(".net") == -1) &&
				((search_mask).search("servicepack") == -1) && 
				((search_mask).search("update") == -1)) {
				progs_array.push(program.item().Name)
				}
			}
			var progs_array_search_sring = progs_array.join("#").toLowerCase().replace(/\s/g, "").replace(/\+/g, "")
			/////////////////////get_from_registry///////////////////
			var RegPath_x32 = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
			var RegPath_x64 = "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
			
			///////////////////////////////arch_check/////////////////
			var archx32 = false
			try {
				var path = WScript.FullName.toLowerCase();
				if ((Arch.toLowerCase().replace(/\s/g, "").replace(/\-/g, "")).search(64) == -1){
					archx32 = true
				}
			} catch(e){}
			
			if (Arch == null || archx32) {
				arch_arr.push(RegPath_x32)
			}else{

					arch_arr.push(RegPath_x32)
					arch_arr.push(RegPath_x64)
			}
			for (cur_arch=0; cur_arch<arch_arr.length; cur_arch++){
				var keyparam = "DisplayName"
				HKLM = 0x80000002;
				var services = GetObject("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\default");
				var Registry = services.Get("StdRegProv"); 

				var Method = Registry.Methods_.Item("EnumKey");
				var p_In = Method.InParameters.SpawnInstance_();
				p_In.hDefKey=HKLM;
				p_In.sSubKeyName = arch_arr[cur_arch];
				var p_Out = Registry.ExecMethod_(Method.Name, p_In);
				var keys=p_Out.sNames.toArray();
				for (i=0; i<keys.length; i++){
					var Method = Registry.Methods_.Item("EnumValues");
					var p_In = Method.InParameters.SpawnInstance_();
					p_In.hDefKey=HKLM;
					p_In.sSubKeyName = arch_arr[cur_arch]+"\\"+keys[i];
					var p_Out = Registry.ExecMethod_(Method.Name, p_In);
					if (p_Out.sNames==null) continue;
					var newkeys=p_Out.sNames.toArray();
					
					for (j=0; j<newkeys.length; j++){
						if (newkeys[j]==keyparam){
							var value_from_reg=wshell.RegRead("HKLM\\"+arch_arr[cur_arch]+"\\"+keys[i]+"\\"+keyparam+"");
							var value_from_reg_for_compare = value_from_reg.toLowerCase().replace(/\s/g, "").replace(/\+/g, "")
							if ((("#"+progs_array_search_sring+"#").search("#"+value_from_reg_for_compare+"#") == -1) &&
							((value_from_reg_for_compare).search("\\(kb") == -1) && 
							((value_from_reg_for_compare).search("visual") == -1) && 
							((value_from_reg_for_compare).search(".net") == -1) && 
							((value_from_reg_for_compare).search("servicepack") == -1) && 
							((value_from_reg_for_compare).search("update") == -1)) {
								progs_array.push(value_from_reg)
							}
						}
					}
				}
			}
			for (var cur_prog_name = 0; cur_prog_name < progs_array.length; cur_prog_name++) {
									programs_inf+="{"
									programs_inf+='"name":"' + progs_array[cur_prog_name] +'"'
									programs_inf+="},"
			}
			programs_inf = programs_inf.slice(0,-1);
			programs_checked = true;
		}
	}catch(e) {error += "SOFT|"}

	///////////////////////////////HDD//////////////////////////
	var HDD_prop_array = [];
	var discs = "";
	var labels = [];
	try {
		if (OS_Type == 1) {
			query_HDD = wmi.ExecQuery("SELECT DeviceID,caption,interfacetype,size FROM Win32_DiskDrive");	
			HDDs = new Enumerator(query_HDD);
			var d = 0;
			for(;!HDDs.atEnd();HDDs.moveNext()){
						var p = 0;
						HDD_prop_array[d]=({'disc':{'name':HDDs.item().caption,'interface':HDDs.item().interfacetype, 'size':(Math.round((HDDs.item().size/1024)/1024)), 'partition':[]}})
						var query = "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='" + HDDs.item().DeviceID + "'} WHERE AssocClass = Win32_DiskDriveToDiskPartition"    
						var DiskPartitions = wmi.ExecQuery(query)
					 
						for (var enumItems2=new Enumerator(DiskPartitions); !enumItems2.atEnd(); enumItems2.moveNext()){
							var wmiDiskPartition = enumItems2.item();
							var wmiLogicalDisks = wmi.ExecQuery("ASSOCIATORS OF {Win32_DiskPartition.DeviceID='" + wmiDiskPartition.DeviceID + "'} WHERE AssocClass = Win32_LogicalDiskToPartition") 
							
							for(var enumItems3=new Enumerator(wmiLogicalDisks ); !enumItems3.atEnd(); enumItems3.moveNext()){
								var wmiLogicalDisk = enumItems3.item();
							   for(HDDs.item().caption in HDD_prop_array) {
								   HDD_prop_array[d].disc.partition[p]={'lable':wmiLogicalDisk.DeviceID, 'size':(Math.round((wmiLogicalDisk.Size/1024)/1024)), 'free':(Math.round((wmiLogicalDisk.FreeSpace/1024)/1024))};
								}
							   p++
							}
						}
				d++
			}
			for (var current_disc = 0; current_disc < HDD_prop_array.length; current_disc++) {
				discs +='{"disc_name":"' + HDD_prop_array[current_disc].disc.name + '","disc_interface":"' + HDD_prop_array[current_disc].disc.interface + '","disc_size":"' + HDD_prop_array[current_disc].disc.size + '","LogicalDiscs":[';
				for (var current_partition = 0; current_partition < HDD_prop_array[current_disc].disc.partition.length; current_partition++) {
					labels[current_partition] = ({'lablename':HDD_prop_array[current_disc].disc.partition[current_partition].lable ,'freespace':HDD_prop_array[current_disc].disc.partition[current_partition].free})
					discs +='{'
					for (var key in HDD_prop_array[current_disc].disc.partition[current_partition])
					{
					discs += '"' + key + '":"' + HDD_prop_array[current_disc].disc.partition[current_partition][key] + '",';
					}
				discs = discs.slice(0,-1);
				discs +='},'
				}
			discs = discs.slice(0,-1);
			discs +=']},'
			}
			discs = discs.slice(0,-1);
		}
	} catch(e) {error += "HDD|"}
	//////////////////////////////////FileSearch///////////////////////////////
   var allfindedfiles = ""
   var filesinprofile = ""
   var configsodinass = ""
   var summaryfilessize = 0
	try {
		//if (1==1) {
		if (!file_checked){
			if (OS_Type == 1){
				var labels_mask = ""
				var files_array = []
				var files_array_in_profile = []
				
				var l = 0;
				for (var current_lable = 0; current_lable < labels.length; current_lable++) {
					labels_mask+= 'Drive = "' +labels[current_lable].lablename+ '" or '
					files_array[l]=({'lable':labels[current_lable].lablename,'extensions':[]})
					files_array_in_profile[l]=({'lable':labels[current_lable].lablename,'extensions':[]})
					var s = 0;
					for (var i = 0; i < ext.length; i++) {
						files_array[l].extensions[i] = ({'type':ext[i],'size':s})
						files_array_in_profile[l].extensions[i] = ({'type':ext[i],'size':s})
					}
					l++
				}
				labels_mask=labels_mask.slice(0,-4);
				var extensions_mask = ""
				for (var i = 0; i < ext.length; i++) {
					extensions_mask+= 'Extension = "'+ext[i]+'" or '
				}
				extensions_mask=extensions_mask.slice(0,-4);
					var profilesearchmask = profile.slice(2).toLowerCase()+"\\";
							
							if (needtocopy){
								f = fso.OpenTextFile(""+profile+"\\copy_files_"+usern+".txt", 2, true, 0);
								f.Close();
								f2 = fso.OpenTextFile(""+profile+"\\copy_files_"+usern+".bat", 2, true, 0);
								f2.Close();
							};
							if (needtoarchive){
								f1 = fso.OpenTextFile(""+profile+"\\archive_files_"+usern+".txt", 2, true, 0);
								f1.Close();
								f3 = fso.OpenTextFile(""+profile+"\\archive_files_"+usern+".bat", 2, true, 0);
								f3.Close();

							}
					
					var TxtFiles = wmi.ExecQuery('Select Path,Drive,Extension,FileName,FileSize FROM CIM_DataFile WHERE ('+labels_mask+')  AND ('+extensions_mask+')');
					var items = new Enumerator(TxtFiles);
					for(;!items.atEnd();items.moveNext()){
						var curfilepath = items.item().Path.toLowerCase().replace(/\s/g, "").replace(/\+/g, "").replace(/\\/g, "#")	
						if (((curfilepath).search("#windows#") == -1) && 
							((curfilepath).search("#programfiles") == -1) && 
							((curfilepath).search("#programdata#") == -1) && 
							((((curfilepath).search("#applicationdata#") >= 0) && (((curfilepath).search("#windowslivemail#") >= 0) || ((curfilepath).search("#outlook") >= 0) || ((curfilepath).search("#1cestart#") >= 0))) || ((curfilepath).search("#applicationdata#") == -1)) &&
							((((curfilepath).search("#appdata#") >= 0) && (((curfilepath).search("mail#") >= 0) || ((curfilepath).search("#outlook") >= 0))) || ((curfilepath).search("#1cestart#") >= 0) || ((curfilepath).search("#appdata#") == -1)) &&
							((curfilepath).search("systemvolumeinformation") == -1) && 
							((curfilepath).search("recycle") == -1)
						) {
							summaryfilessize += +items.item().FileSize
							
							if (needtoarchive){
								f1 = fso.OpenTextFile(""+profile+"\\archive_files_"+usern+".txt", 8, false);
								f1.WriteLine('"'+items.item().Drive+''+items.item().Path+''+items.item().FileName+'.'+items.item().Extension+'"');
								f1.Close();
							};
							
							if (needtocopy){
								f = fso.OpenTextFile(""+profile+"\\copy_files_"+usern+".txt", 8, false, 0);
								f.WriteLine('"'+items.item().Drive+''+items.item().Path+''+items.item().FileName+'.'+items.item().Extension+'" #replace#'+items.item().Path.slice(0,-1)+'"')
								f.Close();
							};
							for (var current_files = 0; current_files < files_array_in_profile.length; current_files++) {
								if (files_array_in_profile[current_files].lable.toLowerCase() == (items.item().Drive)) {
									for (var current_ex = 0; current_ex < files_array_in_profile[current_files].extensions.length; current_ex++) {
										if (files_array_in_profile[current_files].extensions[current_ex].type == (items.item().Extension)) {
											if ((items.item().Path.toLowerCase().indexOf(profilesearchmask)) >= 0) {
												files_array_in_profile[current_files].extensions[current_ex].size += + items.item().FileSize
												if (items.item().Extension == config_1c_ext) {
														var basenamewasfound = false;
														var stream = WScript.CreateObject("ADODB.Stream");
														stream.Charset = 'utf-8';
														stream.Open();
														stream.LoadFromFile(""+items.item().Drive+""+items.item().Path+""+items.item().FileName+"."+items.item().Extension+"");
														while(!stream.EOS){
															var line = stream.ReadText(-2);
															if(basenamewasfound){
																configsodinass+='"Info":"'+line.replace(/\;/g, "").replace(/\"/g, "'")+'"},'
																basenamewasfound = false;
															}
															if(((line.indexOf("\[")) >= 0) || ((line.indexOf("\]")) >= 0)){
																configsodinass+='{"Name":"'+line+'",'
																basenamewasfound = true
															}
														}
														stream.close();
														configsodinass = configsodinass.slice(0,-1);
												}
												
											}else{
												files_array[current_files].extensions[current_ex].size += + items.item().FileSize
											}
										}
									}
								}
							}
						}
					}
					for (var current_files = 0; current_files < files_array_in_profile.length; current_files++) {
						for (var current_ex = 0; current_ex < files_array_in_profile[current_files].extensions.length; current_ex++) {
							if ((Math.round(((files_array_in_profile[current_files].extensions[current_ex].size)/1024)/1024)) > 0) {
								filesinprofile+='{"ext":"'+files_array_in_profile[current_files].extensions[current_ex].type+'",'
								filesinprofile+='"size":"'+Math.round(((files_array_in_profile[current_files].extensions[current_ex].size)/1024)/1024)+'"},'
							}
						}
					}
					filesinprofile = filesinprofile.slice(0,-1);

					for (var current_files = 0; current_files < files_array.length; current_files++) {
						for (var current_ex = 0; current_ex < files_array[current_files].extensions.length; current_ex++) {
							if (Math.round(((files_array[current_files].extensions[current_ex].size)/1024)/1024) > 0) {
								allfindedfiles+='{"lable":"' +files_array[current_files].lable+ '",'
								allfindedfiles+='"ext":"'+files_array[current_files].extensions[current_ex].type+'",'
								allfindedfiles+='"size":"'+Math.round(((files_array[current_files].extensions[current_ex].size)/1024)/1024)+'"},'
							}
						}
					}
					allfindedfiles = allfindedfiles.slice(0,-1);	
									
				if ((needtocopy)||(needtoarchive)){
					copydrive+=':';
					
					function checkfreespaceondrive (arr,lable){
						for (var current_lable = 0; current_lable < arr.length; current_lable++) {
							if (((arr[current_lable].lablename.toLowerCase().indexOf(lable.toLowerCase())) > -1) && ((arr[current_lable].freespace) >= ((Math.round((summaryfilessize/1024)/1024))*2))) {
								enoughspaceondrive = true;
								break
							}else{
								enoughspaceondrive = false;
							}
						}
						return enoughspaceondrive;
					}
					
					
					if (!(checkfreespaceondrive(labels,copydrive))){
						if (autochangecopydrive) {
							var newdiskfinded = false;
							for (var current_lable = 0; current_lable < labels.length; current_lable++) {
								if (checkfreespaceondrive(labels,labels[current_lable].lablename)){
									copydrive = labels[current_lable].lablename
									newdiskfinded = true;
									break
								}
							}
							if (!newdiskfinded){
								WScript.Quit(wshell.popup("No disc finded for copy "+((Math.round((summaryfilessize/1024)/1024))*2)+" Mb. Exit the script"))
							}
						}else{
							WScript.Quit(wshell.popup("Not enough space on selected disc "+copydrive+" (needed space: "+((Math.round((summaryfilessize/1024)/1024))*2)+" Mb) or disc is not available. Exit the script"))
						}
					}

					var copypath = ""+copydrive+"\\"+foldertocopyname+"";
					var archive = ""+copypath+"\\"+archivename+".zip"
					if (!(fso.FolderExists(copypath))){
						fso.CreateFolder(copypath)
					}
				}
				if (needtoarchive){
					if (fso.FileExists(archiver)){
						f = fso.OpenTextFile(""+profile+"\\archive_files_"+usern+".bat", 8, false, 0);
						f.WriteLine(''+archiver+' a -spf -tzip -mx'+compressionlevel+' -mmt4 -scswin -ir@"'+profile+'\\archive_files_'+usern+'.txt" "'+archive+'"')
						f.Close();
						if (autoarchive){
						wshell.run('cmd.exe /k '+archiver+' a -spf -tzip -mx'+compressionlevel+' -mmt4 -scswin -ir@"'+profile+'\\archive_files_'+usern+'.txt" "'+archive+'"')
						}
					}else{
						wshell.popup("Archiver "+archiver+" is not available (check the path)")
					}
				}
				if (needtocopy){
					f = fso.OpenTextFile(""+profile+"\\copy_files_"+usern+".bat", 8, false, 0);
					f.WriteLine('chcp 1251')
					f.WriteLine('@echo off')
					f.WriteLine('for /f "usebackq tokens=*" %%a in ("'+profile+'\\copy_files_'+usern+'.txt") do (')
					f.WriteLine('set line=%%a')
					f.WriteLine('setlocal enabledelayedexpansion')
					f.WriteLine('set "line=echo echo D | xcopy /y /i /q !line:#replace#="'+copypath+'!"')
					f.WriteLine('!line! >> "'+profile+'\\copy_files_'+usern+'.bat"')
					f.WriteLine('endlocal')
					f.WriteLine(')')
					f.WriteLine('@echo on')
					f.WriteLine('rem -------------------------------------------------------------------')
					f.Close();
					
					if (autocopy){
						wshell.run('"'+profile+'\\copy_files_'+usern+'.bat"')
					}
				}
			}else{
				//case of server
					var xpsrv = false;
					var basenamewasfound = false;
					if ((osStr.toLowerCase().replace(/\s/g, "").replace(/(|)/g, "")).search("2003") >= 0){
					xpsrv = true
					}
					var stream = WScript.CreateObject("ADODB.Stream");
					stream.Charset = 'utf-8';
					stream.Open();
					if (xpsrv){
					stream.LoadFromFile(""+profile+"\\Application Data\\1C\\1CEstart\\ibases.v8i");
					}else{
					stream.LoadFromFile(""+profile+"\\AppData\\Roaming\\1C\\1CEStart\\ibases.v8i");
					}
					while(!stream.EOS){
						var line = stream.ReadText(-2);
						if(basenamewasfound){
						configsodinass+='"Info":"'+line.replace(/\;/g, "").replace(/\"/g, "'")+'"},'
						basenamewasfound = false;
						}
						if(((line.indexOf("\[")) >= 0) || ((line.indexOf("\]")) >= 0)){
							configsodinass+='{"Name":"'+line+'",'
							basenamewasfound = true
						}
					}
					stream.close();
					configsodinass = configsodinass.slice(0,-1);									
			}
			try {
				wshell.RegWrite (timekey, ''+ starttime.getTime().toString() +'', "REG_SZ");
			} catch(e) {error += "REGWRITE TIME|"}
			file_finded = true;
		}
	} catch(e) {error += "FILES|"+e.message+""}
	
	/////////////////////////////////1C configs changing////////////////////////////////////
			try{
				if (!odinconfigured){
				var xpsrv = false;
				var xpbackuppath = "\\Application Data\\1C\\1Cv82\\"
				var otherbackuppath1 = "\\AppData\\Roaming\\1C\\1Cv82\\"
				var otherbackuppath2 = "\\AppData\\Local\\1C\\1Cv82\\"
					if ((osStr.toLowerCase().replace(/\s/g, "").replace(/(|)/g, "")).search("xpprofessional") >= 0){
						xpsrv = true
					}
					
					function importcsv (importfile){
						var header_array = [];
						var returnarray = [];
						var linecount = 0;
						var csvstream = WScript.CreateObject("ADODB.stream");
						csvstream.Charset = 'utf-8';
						csvstream.Open();
						csvstream.LoadFromFile(importfile);
						
						while(!csvstream.EOS){
							var line = csvstream.ReadText(-2);
							var sepline = line.split(';')
							for (i=0; i<sepline.length; i++){
								if (linecount == 0){
									header_array.push(sepline[i])
								}else{
									returnarray[linecount-1][header_array[i]]=sepline[i]
								}
							}
							returnarray[linecount] = ({})
							linecount++
						}
						csvstream.close();
						returnarray.pop();
						return returnarray
					}
					
					var bases_arr = importcsv("\\\\10.111.110.3\\1cusers\\list.csv")
					
								/*							
								for (i=0; i<bases_arr.length; i++){
									wshell.popup(""+i+" | "+bases_arr[i].domain_user+" | "+bases_arr[i].path_cfg+" | "+bases_arr[i].base_name+" | "+bases_arr[i].connect_id+" | "+bases_arr[i].srv_1C+"")
								}
								*/
								
				var infoforwrite = ""
				var needtoconfigredirect = false;
				var linecounter = 0
				var newbasenameline = ""
				var objindex = 0
					var outstream = WScript.CreateObject("ADODB.stream");
					outstream.Charset = 'utf-8';
					outstream.Open();
					if (xpsrv){
					outstream.LoadFromFile(""+profile+"\\Application Data\\1C\\1CEStart\\ibases.v8i");	
					}else{
					outstream.LoadFromFile(""+profile+"\\AppData\\Roaming\\1C\\1CEStart\\ibases.v8i");
					}
					//////////////////encoding playing/////////////////
					
					while(!outstream.EOS){
						var line = outstream.ReadText(-2);
						  var lowline = line.toLowerCase()
							if((linecounter == 2)&&((lowline.indexOf("id\=")) >= 0)){
								var idfolder = line.replace(/ID=/g,"")
								
								if (xpsrv){
									fso.CopyFolder(""+profile+""+xpbackuppath+""+idfolder+"", ""+profile+""+xpbackuppath+""+bases_arr[objindex].connect_id+"")
								}else{
									fso.CopyFolder(""+profile+""+otherbackuppath1+""+idfolder+"", ""+profile+""+otherbackuppath1+""+bases_arr[objindex].connect_id+"")
									fso.CopyFolder(""+profile+""+otherbackuppath2+""+idfolder+"", ""+profile+""+otherbackuppath2+""+bases_arr[objindex].connect_id+"")
								}

							}
						
							if((linecounter == 1)&&((lowline.indexOf("connect\=")) >= 0)){
								var needtochangebasename = false;
								for (i=0; i<bases_arr.length; i++){
									if(((lowline.indexOf('"'+bases_arr[i].base_name.toLowerCase()+'"')) >= 0) && 
									(((lowline.indexOf('"'+bases_arr[i].srv_1C.toLowerCase()+'"')) >= 0) || 
									((lowline.indexOf('"'+bases_arr[i].srv_1C.toLowerCase()+':')) >= 0))){
										objindex = i;
										needtochangebasename = true;
										needtoconfigredirect = true;
										break
									}
								}
								if (needtochangebasename){
										infoforwrite += "[Не використовувати ("+newbasenameline.replace(/\[|\]/g, "")+")]"
										infoforwrite += String.fromCharCode(13,10)
										linecounter=2
								}else{
									infoforwrite += newbasenameline
									infoforwrite += String.fromCharCode(13,10)
								}
								newbasenameline = ""
							}
							
							if(((line.indexOf("\[")) >= 0) && ((line.indexOf("\]")) >= 0)){
								newbasenameline = line
								linecounter=1
							}else{
								infoforwrite += line
								infoforwrite += String.fromCharCode(13,10)
							}	
					}
					outstream.close();
							
							var instream = WScript.CreateObject("ADODB.stream");
							instream.Open();
							instream.Type = 2;
							instream.Charset = "utf-8";	
							instream.WriteText(infoforwrite);
							
							if (xpsrv){
								instream.SaveToFile(""+profile+"\\Application Data\\1C\\1CEStart\\ibases.v8i", 2);
							}else{
								instream.SaveToFile(""+profile+"\\AppData\\Roaming\\1C\\1CEStart\\ibases.v8i", 2);
							}
							instream.Close();
							
							if (needtoconfigredirect) {
								var userseachstring = user.toLowerCase().replace(/\s|\-|\\/g, "")
								for (i=0; i<bases_arr.length; i++){
									if((bases_arr[i].domain_user.toLowerCase().replace(/\s|\-|\\/g, "").indexOf(userseachstring)) >= 0){
										var cfgstream = WScript.CreateObject("ADODB.stream");
										cfgstream.Open();
										cfgstream.Type = 2;
										cfgstream.Charset = "UNICODE";	
										cfgstream.WriteText("CommonCfgLocation="+bases_arr[i].path_cfg.toLowerCase()+"");
										if (xpsrv){
											cfgstream.SaveToFile(""+profile+"\\Application Data\\1C\\1CEStart\\1CEStart.cfg", 2);
										}else{
											cfgstream.SaveToFile(""+profile+"\\AppData\\Roaming\\1C\\1CEStart\\1CEStart.cfg", 2);
										}
										cfgstream.Close();
										wshell.RegWrite(odinccheckkey, 1, "REG_DWORD");
										break
									}
								}
							}
				}
			} catch(e) {error += "1C|"+e.message+""}
						
	/////////////////////////////////COPY SAP FILES//////////////////////////////////
	try {
		if ((!file_checked) && (OS_Type == 1)){
			var exeption_arr = [
				"yuzkiv",
				"gtd",
				"gomeniuk",
				"alpatov",
				"konovalchuk",
				"kosygina",
				"levchenko",
				"lehkyi",
				"lytvynenko",
				"miakushko",
				"mostapenko",
				"ponomar",
				"romanenko",
				"ruslan",
				"sydiakov",
				"tocheniyk",
				"tshevchenko"
			];
			
			if (("#"+exeption_arr.join("#").toLowerCase()+"#").search("#"+username+"#") == -1){
				var sapcopypath = ""+profile+"\\AppData\\Roaming\\SAP"
				if (!(fso.FolderExists(sapcopypath))){
					fso.CreateFolder(sapcopypath)
				}
						var scriptpath = "\\\\fs01.ukrtransnafta.com\\Policy"
						/*
						try {
							var strPath = Wscript.ScriptFullName
							var  objFile = fso.GetFile(strPath)
							var scriptpath = fso.GetParentFolderName(objFile) 
						} catch(e) {}
						*/
							fso.CopyFolder(""+scriptpath+"\\SAP\\\*", ""+sapcopypath+"")
			}
		}
	} catch(e) {error += "SAP|"+e.message+""}
	
	/////////////////////////////////////////////////////////////////////////////////
	var allIP = [];
	try {
		query = wmi.ExecQuery("SELECT IPAddress FROM Win32_NetworkAdapterConfiguration");
		items = new Enumerator(query);
		for(;!items.atEnd();items.moveNext()){
			var ipArr = [];
			try {ipArr = items.item().IPAddress.toArray()} catch(e) {}
			for(var i in ipArr) if("0.0.0.0" != ipArr[i]) allIP.push(ipArr[i])
		}
	} catch(e) {error += "IP|"}


	var dameWare = false;
	try {
		query = wmi.ExecQuery("SELECT Caption FROM Win32_Service");
		items = new Enumerator(query);
		for(;!items.atEnd();items.moveNext())
			if(-1 < items.item().Caption.toUpperCase().indexOf("DAMEWARE")){dameWare = true; break}
	} catch(e) {error += "DW|"}

	
	try {
		var key = "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings\\ProxyOverride";
		var value = null;
		try {value = wshell.RegRead(key)} catch(e) {}
		if(null != value){
			if(-1 == value.indexOf("*ukrtransnafta.com"))
				wshell.RegWrite(key, value+";*ukrtransnafta.com;", "REG_SZ")
		}
		else wshell.RegWrite(key, "*ukrtransnafta.com;", "REG_SZ")
	} catch(e) {error += "PROX|"}
		
	var exectime = Math.round((new Date()-starttime)/1000);
	if (error != "") {error=error.slice(0,-1);}
	var json = '{"sAMAccountName":"' + user + '","compname":"' + compname + '","ip":["' + allIP.join('","')+ '"],"osversion":"' + osStr + '","osarch":"' + Arch + '","RAM":"' + RAM + '","CPU_name":"' + CPU_name + '","CPU_freq":"' + CPU_freq + '","GPU_NAME":"' + GPU_NAME + '","GPU_RAM":"' + GPU_RAM + '","GPU_HR":"' + GPU_HR + '","GPU_VR":"' + GPU_VR + '","HDDs":[' + discs + '],"userprofile":"'+ profile +'","filesinprofile":[' + filesinprofile + '],"allfiles":[' + allfindedfiles + '],"printers":[' + printers_inf + '],"bases":[' + configsodinass + '],"programs":[' + programs_inf + '],"note":"'+ note +'","fileschecked":"' + file_finded + '","programschecked":"' + programs_checked + '","dameware":"' + dameWare +'","errors":"' + error + '","execTime":"' + exectime + '"}';
	
	try {
		if (OS_Type == 1) {
			 f = fso.OpenTextFile(""+profile+"\\comp-"+usern+".txt", 2, true, 0);
			 f.Write(""+json+"");
			 f.Close();
			 if ((needtocopy)||(needtoarchive)){
				wshell.popup("All finded files will be copied to "+copypath+"")
			 }
		}
	} catch(e) {}
	/*
	var http = new ActiveXObject("Microsoft.XMLHTTP");
	http.open("POST", "http://invent.ukrtransnafta.com:8088/api/userlogin", false);
	http.setRequestHeader("Host", "invent.ukrtransnafta.com");
	http.setRequestHeader("User-Agent", "Mozilla/4.0 (compatible; Synapse)");
	http.setRequestHeader("Content-Type", "application/json");
	http.send(json);
	*/
	WScript.Echo(json);
	
} catch(e) {};