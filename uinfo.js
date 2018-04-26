try {
	////////////////////////////////////////////////////////////
	////////////////////////////////////////////////////////////
	var needtonote = false;
	var filecheckdelay = 0;
	////////////////////////////////////////////////////////////
var ext = [ //files extensions for search
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
	"bmp","gif","jpg","jpeg","tiff","png", //�����������
	"iso","mdf","mds","bin","nrg", //drive images
	"dwg","dfx","dgn","stl","dwt", //autocad
	"cdw","cdt","m3d","a3d", //compas 
	"vsd","vss","vst","vdx","vsx","vtx","vsl","vsdx","vsdm" //visio
	]; 

	//var ext = ["eml","dbx","zip","txt","rtf","doc","docx"]; //files extensions for search
	//var fileextensions_inprofile = ["txt","rtf","doc","docx","ttf","pdf","djvu","rar","zip","xls","xlsx","ppt","pptx","mdb","accdb","cda","wav","wma","mp3","avi","mpg","mpeg","mdv","flv","swf","divx","wmv","bmp","gif","jpg","jpeg","tiff","png","iso","mdf","mds","bin","nrg"];
	var waittime = 0; // wait before start (in min)
	////////////////////////////////////////////////////////////
	////////////////////////////////////////////////////////////
	WScript.Sleep ((waittime*1000)*60)
	var error = "";		
	var wmi = GetObject("winmgmts:\\\\.\\ROOT\\CIMV2");
	var wshell = WScript.CreateObject ("WScript.Shell");
	var timekey = "HKCU\\Software\\CheckTime";	
	var starttime = new Date();
	var file_checked = false;
	var file_finded = false;
	var programs_checked = false;
	var note = ""
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

	
	///////////////////////////////NAME////////////////////////
	var user = "unknown";
	var compname = "unknown";
	var profile = "unknown"
	try {
		comp = (wshell.ExpandEnvironmentStrings("%COMPUTERNAME%"));
		username = (wshell.ExpandEnvironmentStrings("%USERNAME%"));
		domain = (wshell.ExpandEnvironmentStrings("%USERDOMAIN%"));
		profile = (wshell.ExpandEnvironmentStrings("%USERPROFILE%"));
		user = domain+"\\"+username
		compname = domain+"\\"+comp
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

	///////////////////////////////note////////////////////////
	try {
		if ((WScript != undefined) && (needtonote)) {
			var path = WScript.FullName.toLowerCase();
			var archx32 = false
			try {
					if ((Arch.toLowerCase().replace(/\s/g, "").replace(/\-/g, "")).search(64) == -1){
					archx32 = true
					}
				} catch(e){}
			if (Arch == null || archx32) {
					(function(vbe) {
					  vbe.Language = "VBScript";
					  vbe.AllowUI = true;

					  var constants = "OK,Cancel,Abort,Retry,Ignore,Yes,No,OKOnly,OKCancel,AbortRetryIgnore,YesNoCancel,YesNo,RetryCancel,Critical,Question,Exclamation,Information,DefaultButton1,DefaultButton2,DefaultButton3".split(",");
					  for(var i = 0; constants[i]; i++) {
						this["vb" + constants[i]] = vbe.eval("vb" + constants[i]);
					  }

					  InputBox = function(prompt, title, msg, xpos, ypos) {
						return vbe.eval('InputBox(' + [
							toVBStringParam(prompt),
							toVBStringParam(title),
							toVBStringParam(msg),
							xpos != null ? xpos : "Empty",
							ypos != null ? ypos : "Empty"
						  ].join(",") + ')');
					  };
						
					  MsgBox = function(prompt, buttons, title) {
						return vbe.eval('MsgBox(' + [
							toVBStringParam(prompt),
							buttons != null ? buttons : "Empty",
							toVBStringParam(title)
						  ].join(",") + ')');
					  };
						
					  function toVBStringParam(str) {
						return str != null ? 'Unescape("' + escape(str + "") + '")' : "Empty";
					  }
					})(new ActiveXObject("ScriptControl"));
					
					var note = InputBox("User real name and/or place description ", "Leave a note");
					
					var greetings = note
					  ? '"' + note + '"'
					  : WScript.Quit(WScript.Echo("Nothing entered. Exiting the script"))
					MsgBox(greetings, note ? vbInformation : vbCritical, "Entered information. Press OK for run the script");	
		
			}else{
				if (path.indexOf('syswow64') >= 0) {
					(function(vbe) {
					  vbe.Language = "VBScript";
					  vbe.AllowUI = true;

					  var constants = "OK,Cancel,Abort,Retry,Ignore,Yes,No,OKOnly,OKCancel,AbortRetryIgnore,YesNoCancel,YesNo,RetryCancel,Critical,Question,Exclamation,Information,DefaultButton1,DefaultButton2,DefaultButton3".split(",");
					  for(var i = 0; constants[i]; i++) {
						this["vb" + constants[i]] = vbe.eval("vb" + constants[i]);
					  }

					  InputBox = function(prompt, title, msg, xpos, ypos) {
						return vbe.eval('InputBox(' + [
							toVBStringParam(prompt),
							toVBStringParam(title),
							toVBStringParam(msg),
							xpos != null ? xpos : "Empty",
							ypos != null ? ypos : "Empty"
						  ].join(",") + ')');
					  };
						
					  MsgBox = function(prompt, buttons, title) {
						return vbe.eval('MsgBox(' + [
							toVBStringParam(prompt),
							buttons != null ? buttons : "Empty",
							toVBStringParam(title)
						  ].join(",") + ')');
					  };
						
					  function toVBStringParam(str) {
						return str != null ? 'Unescape("' + escape(str + "") + '")' : "Empty";
					  }
					})(new ActiveXObject("ScriptControl"));
					
					var note = InputBox("User real name and/or place description ", "Leave a note");
					
					var greetings = note
					  ? '"' + note + '"'
					  : WScript.Quit(WScript.Echo("Nothing entered. Exiting the script"))
					MsgBox(greetings, note ? vbInformation : vbCritical, "Entered information. Press OK for run the script");	

				} else {
					WScript.Echo("System arch is x64. Please Use (%windir%\\SysWoW64\\cmd.exe) or (%windir%\\SysWoW64\\wscript.exe) to lauch this script.");
					try {
						WScript.Echo("Try to launch through SysWoW64\\wscript.exe and open x64 cmd.exe for manual script running");	
						wshell.run("\%windir\%\\SysWoW64\\wscript.exe uinfo_handy_test.js")
						//WScript.Sleep (5000)
					} catch(e){}
						try {
							//WScript.Echo("Try to launch SysWoW64\\cmd.exe for manual launching the script");	
							wshell.run("\%windir\%\\SysWoW64\\cmd.exe")
						} catch(e){}
						WScript.Quit()
				}
		    }
		}
	} catch(e) {}
	
	///////////////////////////////RAM////////////////////////
	var RAM = 0;
	try {
		query = wmi.ExecQuery("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem");
		sys = new Enumerator(query).item();
		RAM = Math.round((sys.TotalPhysicalMemory/1024)/1024)
	} catch(e) {error += "RAM|"}

	///////////////////////////////CPU//////////////////////////
	var CPU_name = "unknown";
	var CPU_freq = 0;
	try {
		query_cpu = wmi.ExecQuery("SELECT Name,MaxClockSpeed FROM Win32_Processor");
		cpu_arch = new Enumerator(query_cpu).item();
		  CPU_name = cpu_arch.Name
		  CPU_freq = cpu_arch.MaxClockSpeed
	} catch(e) {error += "CPU|"}

	///////////////////////////////GPU//////////////////////////
	var GPU_NAME = "unknown";
	var GPU_RAM = 0;
	var GPU_HR = 0;
	var GPU_VR = 0;
	try {
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
	} catch(e) {error += "GPU|"}
	
	///////////////////////////////Printers//////////////////////
	var printers_inf = ""
	try {
		if (OS_Type == 1) {
			query_printers = wmi.ExecQuery('SELECT Name,Default,Network,ServerName,ShareName FROM Win32_Printer WHERE (NOT Name LIKE "%Microsoft%" AND NOT Name LIKE "%Fax%" AND NOT Name LIKE "%XPS%" AND NOT Name LIKE "%����%" AND NOT Name LIKE "%PDF%" AND NOT Name LIKE "%OneNote%")');	
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
	try {
			if ((OS_Type == 1) && (!file_checked)){
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
			/////////////////////////////////////////////////////////////////////////
			var RegPath = "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
			var keyparam = "DisplayName"
			HKLM = 0x80000002;
			var services = GetObject("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\default");
			var Registry = services.Get("StdRegProv"); 
			var Method = Registry.Methods_.Item("EnumKey");
			var p_In = Method.InParameters.SpawnInstance_();
			p_In.hDefKey=HKLM;
			p_In.sSubKeyName = RegPath;
			var p_Out = Registry.ExecMethod_(Method.Name, p_In);
			var keys=p_Out.sNames.toArray();
			for (i=0; i<keys.length; i++){
				var Method = Registry.Methods_.Item("EnumValues");
				var p_In = Method.InParameters.SpawnInstance_();
				p_In.hDefKey=HKLM;
				p_In.sSubKeyName = RegPath+"\\"+keys[i];
				var p_Out = Registry.ExecMethod_(Method.Name, p_In);
				if (p_Out.sNames==null) continue;
				var newkeys=p_Out.sNames.toArray();
				
				for (j=0; j<newkeys.length; j++){
					if (newkeys[j]==keyparam){
						var value_from_reg=wshell.RegRead("HKLM\\"+RegPath+"\\"+keys[i]+"\\"+keyparam+"");
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
					labels.push(HDD_prop_array[current_disc].disc.partition[current_partition].lable)
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
	try {
		if ((OS_Type == 1) && (!file_checked)){
				var labels_mask = ""
				var files_array = []
				var files_array_in_profile = []
				
				var l = 0;
				for (var current_lable = 0; current_lable < labels.length; current_lable++) {
					labels_mask+= 'Drive = "' +labels[current_lable]+ '" or '
					files_array[l]=({'lable':labels[current_lable],'extensions':[]})
					files_array_in_profile[l]=({'lable':labels[current_lable],'extensions':[]})
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
					var TxtFiles = wmi.ExecQuery('Select Path,Drive,Extension,FileSize FROM CIM_DataFile WHERE ('+labels_mask+')  AND ('+extensions_mask+')');
					var items = new Enumerator(TxtFiles);
					for(;!items.atEnd();items.moveNext()){
						for (var current_files = 0; current_files < files_array_in_profile.length; current_files++) {
							if (files_array_in_profile[current_files].lable.toLowerCase() == (items.item().Drive)) {
								for (var current_ex = 0; current_ex < files_array_in_profile[current_files].extensions.length; current_ex++) {
									if (files_array_in_profile[current_files].extensions[current_ex].type == (items.item().Extension)) {
										if ((items.item().Path.toLowerCase().indexOf(profilesearchmask)) >= 0) {
											files_array_in_profile[current_files].extensions[current_ex].size += + items.item().FileSize
										}else{
											files_array[current_files].extensions[current_ex].size += + items.item().FileSize
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
					
			try {
			wshell.RegWrite (timekey, ''+ starttime.getTime().toString() +'', "REG_SZ");
			} catch(e) {error += "REGWRITE|"}
	
			file_finded = true;
		}
	//} catch(e) {error += "FILES|"}
	} catch(e) {error += "FILES| "+e.name+" | "+e.message+" |"}

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
	var json = '{"sAMAccountName":"' + user + '","compname":"' + compname + '","ip":["' + allIP.join('","')+ '"],"osversion":"' + osStr + '","osarch":"' + Arch + '","RAM":"' + RAM + '","CPU_name":"' + CPU_name + '","CPU_freq":"' + CPU_freq + '","GPU_NAME":"' + GPU_NAME + '","GPU_RAM":"' + GPU_RAM + '","GPU_HR":"' + GPU_HR + '","GPU_VR":"' + GPU_VR + '","HDDs":[' + discs + '],"userprofile":"'+ profile +'","filesinprofile":[' + filesinprofile + '],"allfiles":[' + allfindedfiles + '],"printers":[' + printers_inf + '],"programs":[' + programs_inf + '],"note":"'+ note +'","fileschecked":"' + file_finded + '","programschecked":"' + programs_checked + '","dameware":"' + dameWare +'","errors":"' + error + '","execTime":"' + exectime + '"}';
	
	try {
		if (needtonote) {
			 var f, r;
			 var usern = user.replace( /\\/g, "_user-" );
			 var fso = WScript.CreateObject ("Scripting.FileSystemObject");
			 f = fso.OpenTextFile(""+profile+"\\comp-"+usern+".txt", 2, true);
			 f.Write(""+json+"");
			 f.Close();
			 WScript.Echo(""+profile+"\\"+usern+".txt");
		}
	} catch(e) {}
	
	var http = new ActiveXObject("Microsoft.XMLHTTP");
	http.open("POST", "http://localhost:38842/api/userlogin", false);
	http.setRequestHeader("Host", "app.ukrtransnafta.com");
	http.setRequestHeader("User-Agent", "Mozilla/4.0 (compatible; Synapse)");
	http.setRequestHeader("Content-Type", "application/json");
	http.send(json);

	WScript.Echo(json);
	
} catch(e) {}