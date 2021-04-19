var oFso = new ActiveXObject("Scripting.FileSystemObject");

var oWShell = new ActiveXObject("WScript.shell");

var connDB = new ActiveXObject("ADODB.Connection");

var rsDB = new ActiveXObject("ADODB.Recordset");    

var objCat = new ActiveXObject("ADOX.Catalog");

var tbl = new ActiveXObject("ADOX.Table");

var CurrMonth = 'December 20';

var TimeVal = '202012'

var gloVar = {

"connStr":["#na",

"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=#mdb#;Jet OLEDB:Database Password=#pwd#;",

"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=#mdb#; Jet OLEDB:Database Password=#pwd#;",

"Driver={Microsoft Access Driver (*.mdb)};Dbq=#mdb#;Uid=Admin;Pwd=#pwd#;",

"Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=#mdb#;Uid=Admin;Pwd=#pwd#;"],

"Curpath":(oWShell.CurrentDirectory).toLowerCase()

}

function pad2(n) { return n < 10 ? '0' + n : n }

function d2folderTime(tmpDT) { try{tmpDT=new Date(tmpDT);} catch(err){tmpDT=new Date();} return ( tmpDT.getFullYear().toString() + pad2(tmpDT.getMonth() + 1) + pad2(tmpDT.getDate()) + pad2(tmpDT.getHours()) + pad2(tmpDT.getMinutes()) + pad2(tmpDT.getSeconds()) ); }

String.prototype.trim=function(){return this.replace(/^\s+|\s+$/g,'');};

Errfile = oFso.CreateTextFile( gloVar.Curpath + '\\ErrFile.log');

var Curpath = oWShell.CurrentDirectory.toLowerCase();

var Inputpath = gloVar.Curpath + '\\input\\';

var Mdbpath = gloVar.Curpath + '\\mstrCalendar.mdb';

var Outputpath = gloVar.Curpath + '\\Output\\';

KillExcel();

tmpstr = '#NA';               

while (tmpstr != "#okok") { try { oFso.copyFile( 'D:\\_bpm_UPD\\db\\mstrCalendar.mdb' , gloVar.Curpath + '\\'); tmpstr = "#okok" ; } catch(err) { tmpstr = err.message ;} }

var filefolder = oFso.GetFolder(Inputpath);

var FileCollection = filefolder.Files;

var FileCount = filefolder.Files.count;

if(FileCount > 0){            

              

                for(var fEnum = new Enumerator(FileCollection); !fEnum.atEnd(); fEnum.moveNext()) {

              

                                KillExcel();

              

                                strFileName = fEnum.item();                                                                                                    

                                var FName= oFso.GetFileName(strFileName);

                                var Fext = oFso.GetExtensionName(strFileName)

                                var Fbase= oFso.GetBaseName(strFileName);

                              

                                var tmpReturnVal = "#okok" ;  

              

                                try {

                                                tmpCnt=1;

                                                while (tmpCnt<gloVar.connStr.length) {

                                                                connDB.connectionString = (gloVar.connStr[tmpCnt]).replace("#mdb#", Mdbpath).replace("#pwd#", "xoru@masalamunch");

                                                                try{connDB.open;}catch(err){oWShell.Popup(err.message )}

                                                                if (connDB.state==1) { gloVar.connStr[0] = gloVar.connStr[tmpCnt]; connDB.close(); tmpCnt=8000; }

                                                                tmpCnt+=1;

                                                              

                                                }

                                }catch(err) { oWShell.popup('1' + err.message)}

                              

                                if (gloVar.connStr[0]=="#na") { tmpReturnVal = "2#err: unable to open .mdb connection."; oWShell.Popup(tmpReturnVal); }                           

                              

                                if(tmpReturnVal =="#okok"){

                                                //try{

                                                                connDB.connectionString = (gloVar.connStr[0]).replace("#mdb#", Mdbpath).replace("#pwd#", "xoru@masalamunch");

                                                                try{connDB.open;} catch(err){ tmpReturnVal = "3#error: " + err.message; oWShell.Popup(tmpReturnVal); }

                                                                              

                                                                if(connDB.state==1) {

                                                                //Cirsql ="UPDATE mstrCalendar SET mstrCalendar.FilePath = Replace([FilePath],'\\10.56.91.11\race\ra\r\_bpm\previous\','\\10.56.91.11\race\bpm\previous\') WHERE (((mstrCalendar.StatusCompliance)<>'Pending' And (mstrCalendar.StatusCompliance)<>'Scheduled'))";

                                                                Cirsql ="UPDATE mstrCalendar SET mstrCalendar.FilePath = iif(InStr([FilePath],'\\\\10.56.91.11\\race\\ra\\r\\_bpm\\previous\\') > 0, 'D:\\bpm\\'& right([FilePath],len([FilePath])-38),[Filepath]) WHERE (((mstrCalendar.StatusCompliance)<>'Pending' And (mstrCalendar.StatusCompliance)<>'Scheduled'))";

                                                              

                                                                //oWShell.popup(Cirsql);

                                                                connDB.Execute(Cirsql);

                                                              

                                                                Cirsql ="DELETE mstrCalendar.Circle  FROM mstrCalendar WHERE (((mstrCalendar.Circle)='TR'))";

                                                                connDB.Execute(Cirsql);

                                                              

                                                              

                                                                                var objExcel = new ActiveXObject ("Excel.Application"); 

                                                                                objExcel.visible = true;

                                                                                objExcel.displayAlerts = false;

                                                                                var objWorkbook = objExcel.Workbooks.open(gloVar.Curpath + '\\input\\' + FName ); 

                                                                              

                                                                                var WKSCount = objWorkbook.Worksheets.count;                        

                                                                                                                                                              

                                                                              

                                                                                for (i=1; i <= WKSCount; i++){

                                                                                                var objWorksheet = objWorkbook.Worksheets(i);

                                                                                                objWorksheet.select;

                                                                                                if(objWorksheet.Name !='Summary'){

                                                                                                objWorksheet.Cells.Clear;

                                                                                                }

                                                                                              

                                                                                }

                                                                              

                                                                                var actcode, shtName

                                                                              

                                                                                if(Fbase.indexOf('_')>0){

                                                                                                actcode = Fbase.substring(0,Fbase.indexOf('_'));

                                                                                                shtName = Fbase.substr(Fbase.indexOf('_')+1);

                                                                              

                                                                                }else{

                                                                                                actcode = Fbase;

                                                                                                shtName = Fbase;

                                                                                }

                                                                              

                                                                                //oWShell.popup(actcode +'\n' + shtName)

                                                                              

                                                                                var objWorksheet = objWorkbook.Worksheets("FileStatus");

                                                                                objWorksheet.select;

                                                                                objWorksheet.Cells.Clear;

                                                                              

                                                                                sql = "SELECT mstrCalendar.CurrMonth, mstrCalendar.ActiCode, mstrCalendar.Circle, mstrCalendar.ReportFileName, mstrCalendar.ConcernPerson02, mstrCalendar.Sch_DT, mstrCalendar.FilePath, mstrCalendar.StatusCompliance FROM mstrCalendar INNER JOIN (SELECT mstrCalendar.ActiCode, mstrCalendar.Circle, Max(mstrCalendar.Proposed_Sch_DT) AS MaxOfSch_DT FROM mstrCalendar WHERE (((mstrCalendar.StatusCompliance)<>'Scheduled') and CurrMonth = '" + CurrMonth +"') GROUP BY mstrCalendar.ActiCode, mstrCalendar.Circle HAVING (((mstrCalendar.ActiCode)='"+ actcode +"')) )  AS MaxTab ON (mstrCalendar.Proposed_Sch_DT = MaxTab.MaxOfSch_DT) AND (mstrCalendar.Circle = MaxTab.Circle) AND (mstrCalendar.ActiCode = MaxTab.ActiCode) GROUP BY mstrCalendar.CurrMonth, mstrCalendar.ActiCode, mstrCalendar.Circle, mstrCalendar.ReportFileName, mstrCalendar.ConcernPerson02, mstrCalendar.Sch_DT, mstrCalendar.FilePath, mstrCalendar.StatusCompliance"

                                                                                rsDB.open(sql, connDB, 1, 3);

                                                                              

                                                                                if(rsDB.eof==false) {

                                                                                              

                                                                                    var fieldcount = rsDB.Fields.count;                                     

                                                                                                var iCol=1

                                                                                                while (iCol<=fieldcount){

                                                                                                                objWorksheet.cells(1, iCol).value =  rsDB.fields(iCol - 1).name

                                                                                                                iCol=iCol+1

                                                                                                }

                                                                                                              

                                                                                                objWorksheet.cells(2, 1).copyFromRecordset(rsDB)

                                                                                              

                                                                                                if(rsDB.eof==true) { rsDB.movefirst;}                                                   

                                                                                              

                                                                                                // for each  Copy file to /input/temp

                                                                                                                                                                              

                                                                                                while (rsDB.eof==false && tmpReturnVal =="#okok") {

                                                                                              

                                                                                                                //oWShell.Popup(rsDB.fields('FilePath'));

                                                                                                              

                                                                                                                if(            rsDB.fields('FilePath').value != null ) {

                                                                                                              

                                                                                                                tmpstr = '#NA';               

                                                                                                                while (tmpstr != "#okok") { try { oFso.copyFile( rsDB.fields('FilePath').value , gloVar.Curpath + '\\input\\Temp\\'); tmpstr = "#okok" ; } catch(err) { tmpstr = err.message ;} }

                                                                                                              

                                                                                                                //oWShell.popup(rsDB.fields('FilePath').value );

                                                                                                              

                                                                                                                objWorkbook.activate;

                                                                                                                objWorksheet = objWorkbook.Worksheets(rsDB.fields('Circle').value);

                                                                                                                objWorksheet.select

                                                                                                                objWorksheet.Cells.Clear;

                                                                                                                objWorksheet.Range("A1").select

                                                                                                              

                                                                                                                var OPFileName= oFso.GetFileName(rsDB.fields('FilePath').value);

                                                                                                                try { var WBookOPCOFile = objExcel.Workbooks.open(gloVar.Curpath + '\\input\\Temp\\' + OPFileName );  } catch(err){ rsDB.MoveNext(); continue; }

                                                                                                                try{ var WksheetOPCOFile = WBookOPCOFile.Worksheets(shtName); } catch(err){var WksheetOPCOFile = WBookOPCOFile.Worksheets(1);}

                                                                                                                try{WksheetOPCOFile.select;}catch(err){ WBookOPCOFile.Worksheets(1).select; }

                                                                                                                objExcel.ActiveWindow.FreezePanes = false;

                                                                                                                WksheetOPCOFile.Cells.copy;

                                                                                                              

                                                                                                                //objWorksheet.Paste;

                                                                                                                objWorksheet.Range("A1").PasteSpecial();

                                                                                                                objWorkbook.activate;

                                                                                                                objExcel.Run('ActSheetPasteValue');

                                                                                                              

                                                                                                              

                                                                                                              

                                                                                                                WBookOPCOFile.close();

                                                                                                                }

                                                                                                              

                                                                                                                rsDB.MoveNext();

                                                                                                              

                                                                                                }                                                                                            

                                                                                              

                                                                                }else{

                                                                                                objWorksheet.Range("A1").value =  "No data received" ;

                                                                                              

                                                                                }

                                                                              

                                                                                                try{rsDB.close();}catch(err){}

                                                                                                objWorkbook.save;

                                                                                              

                                                                                              

                                                                                                tmpstr = '#NA';

                                                                                                var FileName = TimeVal + '_Summary_' + Fbase + '.xlsx';

                                                                                                FileName = gloVar.Curpath + '\\Output\\' + TimeVal + '\\' + FileName                                                                                                                                              

                                                                                                if(oFso.folderExistS(gloVar.Curpath + '\\Output\\' + TimeVal ) == false) {oFso.CreateFolder( gloVar.Curpath + '\\Output\\' + TimeVal )}

                                                                                              

                                                                                                objWorkbook.saveas(FileName, 51);                                                                                                                                    

                                                                                                objWorkbook.close();

                                                                                                objExcel.quit();                                                                                                                               

                                                                                              

                                                                                              

                                                                                                //while (tmpstr != "#okok") { try { oFso.copyFile( gloVar.Curpath + '\\input\\' + FName , gloVar.Curpath + '\\Output\\' + FileName); tmpstr = "#okok" ; } catch(err) { tmpstr = err.message ;} }

                                                                                              

                                                                                                oFso.deleteFile(gloVar.Curpath + '\\input\\Temp\\*.*');

                                                                              

                                                                              

                                                                              

                                                                                //Copy from recordset to Excel sheet.

                                                              

                                                                }

                                                                try{connDB.close();} catch(err){}

                                                              

                                                //} catch(err){oWShell.popup('last' + err.message)}

                                                try{ oFso.MoveFile(gloVar.Curpath + '\\input\\' + FName , gloVar.Curpath + '\\input\\Done\\'); } catch(err){}

                                              

                                                //oWShell.popup('Done');

                                              

                                }

                              

                }            

}                            

function KillExcel(){

                var objWMIService, objProcess, colProcess, strComputer, strList

                strComputer = "."

                processName = "EXCEL.EXE";

              

                var objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\\\" + strComputer + "\\root\\cimv2")

                var colProcess = objWMIService.ExecQuery ("Select * from Win32_Process where CommandLine Like '%"+ processName +"%'")

                listcount = 0

                for(var objEnum = new Enumerator(colProcess); !objEnum.atEnd(); objEnum.moveNext()) {

                                objEnum.item().Terminate()

                }

                //oWShell.popup('Excel killed');

              

}            