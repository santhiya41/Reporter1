package CommonLibraries;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.apache.commons.io.IOUtils;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFHyperlink;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;

import com.google.common.io.Files;

import cucumber.api.java.sl.Ce;

public class FunctionLibraries {

	//Get the current date
		public static String fn_GetDate(){
			//Get the current system date
			String Current_Date = new SimpleDateFormat("M-d-yyyy").format(Calendar.getInstance().getTime());		
			return Current_Date;
		}
		
		//Get the time
		public static String fn_GetTime(){
			//Get the current system date
			String Current_Time = new SimpleDateFormat("HHmmss").format(Calendar.getInstance().getTime());		
			return Current_Time;
		
		}
		
	//Create Folder with current date
		public static File fn_CreateResultFolder() throws Exception{

			//Get the date
			String current_Date = fn_GetDate();

			String Resultpath = AppConstant.RESULT_FOLDERLOCATION+current_Date;
			File file = new File(AppConstant.RESULT_FOLDERLOCATION);
			File file1 = new File(Resultpath);
			String[] Filenames = file.list();
			boolean Folderexist = false;
			
			//Verify the result folder exists 
			for(String name : Filenames)
			{
				if (name.toString().equals(current_Date.toString())) {
					Folderexist = true;
					break;
				}
			}
			
			//Create the result folder if it doesn't exists
			if (!Folderexist) {
				file1.mkdir();
				Folderexist = true;
			}
			
			return file1;
		}
		
		//Creates Scenario folder
		public static File fn_CreateFeatureFolder(File resultFoldername, String FeatureName) {
			
			String FeaturePath = resultFoldername+"\\"+FeatureName;
			File FeatureNameFolderpath = new File(FeaturePath);
			String[] FeatureNames = resultFoldername.list();
			boolean FeatureFolderexist = false;
			
			//Verify the result folder exists 
			for(String feature : FeatureNames)
			{
				if (feature.toString().equals(FeatureName)) {
				     FeatureFolderexist = true;
					break;
				}
			}
			
			//Create the folder with the fund Name
			if (!FeatureFolderexist) {
				FeatureNameFolderpath.mkdir();
			}
			return FeatureNameFolderpath;
		}
		
		//Create the test script name folder
		public static File fn_CreateTestScriptNameFolder(File featureNameFolder,String testscirptname) {
			
			String TestscriptfolderName = featureNameFolder+"\\"+testscirptname+"_"+fn_GetDate()+"_"+fn_GetTime();
			String SnapshotFoldername =  TestscriptfolderName+"\\Snapshot";
			File TestScriptFolderpath = new File(TestscriptfolderName);
			File SnapshotFolderpath = new File(SnapshotFoldername);
			
			//String[] TestScriptFolder = fundNameFolder.list();
			//boolean FundFolderexist = false;
			
			////Verify the result folder exists 
			//for(String TestFolder : TestScriptFolder)
			//{
				//if (TestFolder.toString().equals(testscirptname)) {
					//FundFolderexist = true;
					//break;
				//}
			//}
			
			//Create the folder with the test name
			//if (!FundFolderexist) {
				TestScriptFolderpath.mkdir();
				SnapshotFolderpath.mkdir();
			//}
				
			return TestScriptFolderpath;
		}
		
		//Creates the HTML file
		public static File fn_CreateHTML(File RuntimeResultFolderlocation) throws Exception {
			
			String ResFold = RuntimeResultFolderlocation+"\\ResultTemplate.html";
			System.out.println(ResFold);
			File ToHTMLfolder = new File(ResFold);
			File FromHTMLfolder = new File(AppConstant.HTML_LOCATION);
			Files.copy(FromHTMLfolder, ToHTMLfolder);	
			return ToHTMLfolder;
			
		}
		
		//Open the TR tag in the HTML
		public static void fn_Open_TR_Tag(String HTML_FilePath) throws IOException {
			
			FileWriter filewrite = new FileWriter(HTML_FilePath,true);
		    BufferedWriter BufferWrite = new BufferedWriter(filewrite);
		    PrintWriter write = new PrintWriter(BufferWrite);
		    
		    write.println("<tr>");
		    write.close();
		    BufferWrite.close();
		    filewrite.close();
		}
		
		//Close the TR tag in the HTML
		public static void fn_Close_TR_Tag(String HTML_FilePath) throws IOException {
			
			FileWriter filewrite = new FileWriter(HTML_FilePath,true);
		    BufferedWriter BufferWrite = new BufferedWriter(filewrite);
		    PrintWriter write = new PrintWriter(BufferWrite);
		    
		    write.println("</tr>");
		    write.close();
		    BufferWrite.close();
		    filewrite.close();
		}
		
		//Writes a new entry in the HTML reporter by using the existing screenshot path
		public static String fn_Update_HTML(String HTML_FilePath, String TestCase,String Status, String Step, String Description, WebDriver Driver ,String Snapshot_Path) throws IOException, UnsupportedOperationException, Throwable {
			
			//Open TR Tag
			fn_Open_TR_Tag(HTML_FilePath);
		    
			FileWriter filewrite = new FileWriter(HTML_FilePath,true);
		    BufferedWriter BufferWrite = new BufferedWriter(filewrite);
		    PrintWriter write = new PrintWriter(BufferWrite);
		    
		    String Time = new SimpleDateFormat("HH:mm:ss").format(Calendar.getInstance().getTime());
		    String Snapshotvalue = null;
		    String DataToAppend = null;
		    
		    //Capture the snapshot
			Snapshotvalue = "<a href = "+Snapshot_Path+">Snap Shot</a>";
		    
		    //populate the appending line
		    if (Status.equalsIgnoreCase("PASS")) {
		    	DataToAppend = "<td>"+TestCase+"</td><td><font color=\"limegreen\">PASS</font></td><td>"+Step+"</td><td>"+Description+"</td><td>"+Time+"</td><td>"+Snapshotvalue+"</td>";
			} else if (Status.equalsIgnoreCase("FAIL")) {
				DataToAppend = "<td>"+TestCase+"</td><td><font color=\"Red\">FAIL</font></td><td>"+Step+"</td><td>"+Description+"</td><td>"+Time+"</td><td>"+Snapshotvalue+"</td>";
			} else if (Status.equalsIgnoreCase("WARN")) {
				DataToAppend = "<td>"+TestCase+"</td><td><font color=\"Yellow\">WARN</font></td><td>"+Step+"</td><td>"+Description+"</td><td>"+Time+"</td><td>"+Snapshotvalue+"</td>";
			}	

		    write.println(DataToAppend);
		    write.close();
		    BufferWrite.close();
		    filewrite.close();
			
		    //Close TR Tag
		    fn_Close_TR_Tag(HTML_FilePath);
		    
	    	return Snapshot_Path;
			
		}
		
		//Writes a new entry in the HTML reporter by taking the screenshot
		public static String fn_Update_HTML(String HTML_FilePath, String TestCase,String Status, String Step, String Description, WebDriver Driver ,boolean snapshot) throws IOException, UnsupportedOperationException, Throwable {
			
			//Open TR Tag
			fn_Open_TR_Tag(HTML_FilePath);
		    
			FileWriter filewrite = new FileWriter(HTML_FilePath,true);
		    BufferedWriter BufferWrite = new BufferedWriter(filewrite);
		    PrintWriter write = new PrintWriter(BufferWrite);
		    
		    String Time = new SimpleDateFormat("HH:mm:ss").format(Calendar.getInstance().getTime());
		    String Snapshotvalue = null;
		    String DataToAppend = null;
		    String Screenshotname = "Snapshots_"+fn_GetDate()+"_"+fn_GetTime()+".PNG";
		    String Snapshotpath = HTML_FilePath.replace("ResultTemplate.html", "")+"Snapshot\\"+Screenshotname;
		    
		    //Capture the snapshot
		    if (snapshot) {
		    	//BufferedImage screencapture = new Robot().createScreenCapture(new Rectangle(Toolkit.getDefaultToolkit().getScreenSize()));
		    	//File file = new File(Snapshotpath);
		    	//Save as JPEG
		    	//ImageIO.write(screencapture,"jpg",file);
		    	
				File screen = ((TakesScreenshot)Driver).getScreenshotAs(OutputType.FILE);
				File ScreeenshotLocation = new File(Snapshotpath);
				org.apache.commons.io.FileUtils.copyFile(screen,ScreeenshotLocation);
				Snapshotvalue = "<a href = "+Snapshotpath+">Snap Shot</a>";
		    	
			} else {
				Snapshotvalue = "NA";
			}
		    
		    //populate the appending line
		    if (Status.equalsIgnoreCase("PASS")) {
		    	DataToAppend = "<td>"+TestCase+"</td><td><font color=\"limegreen\">PASS</font></td><td>"+Step+"</td><td>"+Description+"</td><td>"+Time+"</td><td>"+Snapshotvalue+"</td>";
			} else if (Status.equalsIgnoreCase("FAIL")) {
				DataToAppend = "<td>"+TestCase+"</td><td><font color=\"Red\">FAIL</font></td><td>"+Step+"</td><td>"+Description+"</td><td>"+Time+"</td><td>"+Snapshotvalue+"</td>";
			} else if (Status.equalsIgnoreCase("WARN")) {
				DataToAppend = "<td>"+TestCase+"</td><td><font color=\"Yellow\">WARN</font></td><td>"+Step+"</td><td>"+Description+"</td><td>"+Time+"</td><td>"+Snapshotvalue+"</td>";
			}	

		    write.println(DataToAppend);
		    write.close();
		    BufferWrite.close();
		    filewrite.close();
			
		    //Close TR Tag
		    fn_Close_TR_Tag(HTML_FilePath);
		    
		    //Return the snapshot path
		    if (snapshot) {
		    	return Snapshotpath;
			} else {
				return "";
			}
		}
		    
		  //End the HTML reporter by calculating the number of PASS/FAIL/WARNING
			public static void fn_End_HTML (String HTMLFilePath) {
				
				int beginIndex = 0;
				int endIndex = 0;
				String Executionstart = "EXECUTION STARTED ON </td><td>";
				int Extstart = Executionstart.trim().length();
				
				try {
					
					FileInputStream fileinput = new FileInputStream(HTMLFilePath);
					String content = IOUtils.toString(fileinput, "UTF-8");
					beginIndex = content.indexOf(Executionstart);
					endIndex = beginIndex+Extstart+8;
					
					int passcount = 0;
					int failcount = 0;
					int warncount = 0;
					
					//Calculate the number of Pass, Fail, Warnings		
					if (content.contains(">PASS<")) {
						String[] PassRepetition = content.split(">PASS<");
						passcount = PassRepetition.length;
						passcount = passcount - 1;
					} if (content.contains(">FAIL<")) {
						String[] FailRepetition = content.split(">FAIL<");
						failcount = FailRepetition.length;
						failcount = failcount - 1;
					} if (content.contains(">WARN<")) {
						String[] WarnRepetition = content.split(">WARN<");
						warncount = WarnRepetition.length;
						warncount = warncount - 1;
					}
					
					String passcnt = String.valueOf(passcount);
					String failcnt = String.valueOf(failcount);
					String warncnt = String.valueOf(warncount);
					
					//Convert the start time to the milliseconds
					String StartTime = content.substring(beginIndex+Extstart, endIndex);
					String[] SplitTime = StartTime.split(":");
					long longstarthour = Long.parseLong(SplitTime[0])*60*60*1000;
					long longstartminute = Long.parseLong(SplitTime[1])*60*1000;
					long longstartsecond = Long.parseLong(SplitTime[2])*1000;
					long StartTimemillisecond = longstarthour+longstartminute+longstartsecond;
					
					//Convert the end time to the milliseconds
					String Current_Time = new SimpleDateFormat("HH:mm:ss").format(Calendar.getInstance().getTime());
					String[] SplitCurrentime = Current_Time.split(":");
					long longendhour = Long.parseLong(SplitCurrentime[0])*60*60*1000;
					long longendminute = Long.parseLong(SplitCurrentime[1])*60*1000;
					long longendsecond = Long.parseLong(SplitCurrentime[2])*1000;
					long EndTimemillisecond = longendhour+longendminute+longendsecond;
					long timedifference = EndTimemillisecond - StartTimemillisecond;
					
					//Convert the time difference to hh:mm:ss format
					String hms = String.format("%02d:%02d:%02d", TimeUnit.MILLISECONDS.toHours(timedifference),TimeUnit.MILLISECONDS.toMinutes(timedifference) % TimeUnit.HOURS.toMinutes(1),TimeUnit.MILLISECONDS.toSeconds(timedifference) % TimeUnit.MINUTES.toSeconds(1));
					String[] splittimeduration = hms.split(":");
					String Timeduration = splittimeduration[0]+" hr:"+splittimeduration[1]+" min:"+splittimeduration[2]+" sec";
					
					//Update the html reporter
					content = content.replaceAll("KEY_END_TIME", Current_Time);
					content = content.replaceAll("KEY_DURATION_TIME", Timeduration);
					content = content.replaceAll("KEY_PASS", passcnt);
					content = content.replaceAll("KEY_FAIL", failcnt);
					content = content.replaceAll("KEY_WARNING", warncnt);
					
					FileOutputStream fileoutput = new FileOutputStream(HTMLFilePath);
					IOUtils.write(content,fileoutput , "UTF-8");
					fileinput.close();
					fileoutput.close();
					
					//Add the filter
					fn_addFilter(HTMLFilePath);
					
				} catch (Exception e) {
				
				}
						
			}
				
			//Start the HTML by updating the table
			public static void fn_Start_HTML(String HTMLFilePath, String test_name, String feature_name) throws IOException {
				
				String MonthNumber = new SimpleDateFormat("MMM").format(Calendar.getInstance().getTime());
				String Date = new SimpleDateFormat("d").format(Calendar.getInstance().getTime());
				String Year = new SimpleDateFormat("yyyy").format(Calendar.getInstance().getTime());
				String Current_Time = new SimpleDateFormat("HH:mm:ss").format(Calendar.getInstance().getTime());
				String ExecutionStarte_date = Date+"/"+MonthNumber+"/"+Year;
				String ExecutionStart_Time = Current_Time;
				
				FileInputStream fileinput = new FileInputStream(HTMLFilePath);
				String content = IOUtils.toString(fileinput, "UTF-8");
				content = content.replaceAll("KEY_WORKFLOW_NAME", test_name);
				content = content.replaceAll("KEY_START_TIME", ExecutionStart_Time);
				content = content.replaceAll("KEY_EXECUTIONDATE", ExecutionStarte_date);
				content = content.replaceAll("KEY_FUND_NAME", feature_name);
				FileOutputStream fileoutput = new FileOutputStream(HTMLFilePath);
				IOUtils.write(content,fileoutput , "UTF-8");
				fileinput.close();
				fileoutput.close();
			}
			
			//Add the filter to the HTML reporter
			public static void fn_addFilter(String filePath)
			{
				try
				{
					PrintWriter out = new PrintWriter(new BufferedWriter(new FileWriter(filePath, true)));
				    out.println("<script language=\"javascript\" type=\"text/javascript\">");
				    out.println("//<![CDATA[");
				    out.println("setFilterGrid(\"table1\");");
				    out.println("//]]>");
				    out.println("</script>");
				    out.close();
				}
				catch(IOException e)
				{
					
				}
			}
			
			public static void fn_batch_updateHTML(){
			ArrayList<String> Fund_Names = new ArrayList<String>();
			ArrayList<String> Script_Names = new ArrayList<String>();
			ArrayList<String> ETF_Ticker = new ArrayList<String>();
			ArrayList<String> Component_Tab = new ArrayList<String>();
			ArrayList<String> HTML_Path = new ArrayList<String>();
			ArrayList<String> Status = new ArrayList<String>();
			ArrayList<String> AutomationExecTime = new ArrayList<String>();
			ArrayList<String> AutomationExecStartTime = new ArrayList<String>();
			HashMap<String, String> ExecutionFlag = new HashMap<String, String>();
			
			Fund_Names.clear();
			Script_Names.clear();
			ETF_Ticker.clear();
			Component_Tab.clear();
			Status.clear();
			ExecutionFlag.clear();
			
			String Date_ResultFolder = null;
			String Final_ResultFolder = null;
			String HTML_ResultFile = "ResultTemplate.html";	
			
			try {
					
				String BatchReportFile = AppConstant.BatchReportTemplatePath.replaceAll("Template.xls", "_"+fn_GetDate()+".xls");
				String Final_BatchReportFile = BatchReportFile.replaceAll("Settings", "");
				File ToBatchRpoertFile = new File(Final_BatchReportFile);
				File FromBatchRpoertFile = new File(AppConstant.BatchReportTemplatePath);
				File ReporterFileCopy = new File(ToBatchRpoertFile.getPath().replaceAll("BatchExecutionReport", "BatchExecutionReport//BatchReporterCopy//"));
				Files.copy(FromBatchRpoertFile, ToBatchRpoertFile);	
				
				//Get the File Name
				AppConstant.BatchReporterExcel = ToBatchRpoertFile.getPath();
				
				//Get the Date Folder in the result location
				Date_ResultFolder = fn_GetDate();
				
				//Get the Final Result Folder
				Final_ResultFolder = AppConstant.RESULT_FOLDERLOCATION+Date_ResultFolder+"\\";
				
				//Get the Fund Names
				Fund_Names = fn_Get_Fund_Names();
				
				//Get the Script Names
				Script_Names = fn_Get_Script_Names();
				
				//Get the Execution Flag
				//ExecutionFlag = fn_Get_ExecutionFlag(Fund_Names, Script_Names);
				int totalPassCount = 0;
				int totalFailCount = 0;
					
					//Iterate with the script name
					for (int script = 0; script < Script_Names.size(); script++) {
						
						String fnd_nme = Fund_Names.get(script);
						
						String Scrpt_nme = Script_Names.get(script);
						String Final_DataFile_Path = null;
						
					//	if (ExecutionFlag.get(fnd_nme+"|"+Scrpt_nme).equalsIgnoreCase("Y")) {
							
							String LatestModifiedFolder = null;
							String LatestHTMLPath = null;
							
							//Get the Latest Modified Folder
							LatestModifiedFolder = fn_LastModified_Directory(Final_ResultFolder+fnd_nme,Scrpt_nme);
							
							if (LatestModifiedFolder != null) {
								
								String TestStatus = null; 
								String ExecutionTime = null;
								String ExecutionStartTime = null;
								
								
								//Get the latest HTML path
								LatestHTMLPath = LatestModifiedFolder+"\\"+HTML_ResultFile;
															
								//Get the Pass/Fail Status from the HTML Reporter
								TestStatus = fn_Get_Test_Status(LatestHTMLPath);
								
								ExecutionTime = fn_Get_Test_ExecutionTime (LatestHTMLPath);
								
								ExecutionStartTime = fn_Get_Test_ExecutionStartTime (LatestHTMLPath);
								
								//Get the Data File Path Name
								Final_DataFile_Path = AppConstant.OuputDatalocation+fnd_nme+"\\"+Scrpt_nme+".xls";
								
								//Get the Failed Data Point Level Status
								//fn_Get_Fail_DataPoint_Entry(Final_DataFile_Path,fnd_nme);
								
								ETF_Ticker.add(fnd_nme);
								Component_Tab.add(Scrpt_nme);
								HTML_Path.add(LatestHTMLPath);
								Status.add(TestStatus);
								AutomationExecTime.add(ExecutionTime);
								AutomationExecStartTime.add(ExecutionStartTime);
								totalPassCount = totalPassCount + fn_Get_Test_ExecutionPassCount(LatestHTMLPath);
								totalFailCount = totalFailCount + fn_Get_Test_ExecutionFailCount(LatestHTMLPath);
							//}											
					}
					
				}	
				
				//Update the reporter
				fn_Update_Reporter1(Script_Names, Status, HTML_Path, AutomationExecTime, AutomationExecStartTime.get(0), totalPassCount, totalFailCount);
				//Run the VBS file to send the mail
				Runtime.getRuntime().exec("cscript "+AppConstant.VBS_RESULT_FilePath);
			} catch (Throwable e) {
				
				e.printStackTrace();
				
			}
		  
	}
			
			//Get the fund names in ArrayList
			public static ArrayList<String> fn_Get_Fund_Names() throws Throwable {
				
				ArrayList<String> FundName = new ArrayList<String>();
						
				try {
					
					FileInputStream Fip = new FileInputStream(AppConstant.MasterInputExcel);
					HSSFWorkbook wrkbk = new HSSFWorkbook(Fip);
					HSSFSheet Sheet = wrkbk.getSheet("Script_Execution");
					
					Iterator<Row> Row = Sheet.rowIterator();
					
					Row.next();
					
					while (Row.hasNext()) {
						
						HSSFRow row = (HSSFRow) Row.next();
						
						HSSFCell cell = row.getCell(0);
						
						FundName.add(cell.getStringCellValue());
						
					}
					
					//Close the workbook
					wrkbk.close();
					Fip.close();
					
				} catch (FileNotFoundException e) {
					
					e.printStackTrace();
					
				}
				
				return FundName;
				
			}
			
			//Get the Script Names
			public static ArrayList<String> fn_Get_Script_Names() throws Throwable {
			
				ArrayList<String> ScriptName = new ArrayList<String>();
				
				try {
					
					FileInputStream Fip = new FileInputStream(AppConstant.MasterInputExcel);
					HSSFWorkbook wrkbk = new HSSFWorkbook(Fip);
					HSSFSheet Sheet = wrkbk.getSheet("Script_Execution");
					
					Iterator<Row> Row = Sheet.rowIterator();
					
					Row.next();
						
						while (Row.hasNext()) {
							
							HSSFRow row = (HSSFRow) Row.next();
							
							HSSFCell cell = row.getCell(1);
												
							ScriptName.add(cell.getStringCellValue());
						
					}
					
					//Close the workbook
					wrkbk.close();
					Fip.close();
					
				} catch (FileNotFoundException e) {
					e.printStackTrace();
				}
				
				return ScriptName;
				
			}
			
			//Get the Execution Flag for the Funds and Scripts
			public static HashMap<String,String> fn_Get_ExecutionFlag(ArrayList<String> FundName, ArrayList<String> ScriptName) throws Throwable {
				
				HashMap<String, String> ExecutionFlag = new HashMap<String, String>();
				
				try {
					
					FileInputStream Fip = new FileInputStream(AppConstant.MasterInputExcel);
					HSSFWorkbook Wrkbk = new HSSFWorkbook(Fip);
					HSSFSheet Sheet = Wrkbk.getSheet("Script_Execution");
					
					for (int Rownum = 0; Rownum < FundName.size(); Rownum++) {
						
						String Fund_Name = null;
						
						//Get the Fund Name
						Fund_Name = FundName.get(Rownum);
						
						HSSFRow Row = (HSSFRow) Sheet.getRow(Rownum+1);
						
						for (int Colnum = 0; Colnum < ScriptName.size(); Colnum++) {
							
							String Script_Name = null;
							String Key = null;
							String Value = null;
							
							//Get the Script Name
							Script_Name = ScriptName.get(Colnum);
							
							HSSFCell Cell = Row.getCell(Colnum+1);
							
							Key = Fund_Name+"|"+Script_Name;
							Value = Cell.getStringCellValue().toString().trim();
							
							//Update the Hash Map
							ExecutionFlag.put(Key,Value);
							
						}
						
					}
					
					//Close the workbook
					Wrkbk.close();
					
				} catch (FileNotFoundException e) {
					e.printStackTrace();
				}
				
				return ExecutionFlag;
			
			}
			
			//Get the last modified folder path
			public static String fn_LastModified_Directory(String FolderPath, String TestCaseName) {
				
				ArrayList<Long> lastmodified = new ArrayList<Long>();
				HashMap<Long,String> File_LastModified = new HashMap<Long, String>();
				long LatestModified = 0;
				String FinalValue = null;
				
				try {
					
					File directory = new File(FolderPath);
					File[] subdirs = directory.listFiles();
					
					//Get the Last Modified date to the array list
					for(File subdir : subdirs){
						if (subdir.getPath().toString().contains(TestCaseName)) {
							File file = new File(subdir.getPath());
							lastmodified.add(file.lastModified());
							File_LastModified.put(file.lastModified(),subdir.getPath());
						}
						
					}
					
					//Sort the array List
					Collections.sort(lastmodified);
					
					//Get the last value in the array list
					if (lastmodified.size() == 0) {
						LatestModified = 0;
					}else if (lastmodified.size() == 1) {
						LatestModified = lastmodified.get(0);
					} else {
						LatestModified = lastmodified.get(((lastmodified.size()) - 1));
					}
					
					//Get the file path of the last modified folder
					FinalValue = File_LastModified.get(LatestModified);
					
				} catch (Exception e) {
				
				}
				
				return FinalValue;
				
			}	

			//Get the HTML reporter status
			public static String fn_Get_Test_Status(String HTMLFilePath) throws Exception {
				
				String content = null;
				String SplitContent = null;
				String FinalStatus = "Fail";
				
				try {
					
					FileInputStream fileinput = new FileInputStream(HTMLFilePath);
					content = IOUtils.toString(fileinput, "UTF-8");
					SplitContent = ">FAIL </font></td><td class=\"2\">";
					String[] SplitValue = content.split(SplitContent);
					
					//Find the first occurrence of the string '<'
					int position = SplitValue[1].indexOf("<", 0);
					
					try {
						
						//Get the Fail Count
						int FailCount = Integer.valueOf(SplitValue[1].substring(0, position));
						
						if (FailCount > 0) {
							FinalStatus = "Failed";
						} else if (FailCount == 0) {
							FinalStatus = "Passed";
						}
						
					} catch (Exception e) {
						
						e.printStackTrace();
						FinalStatus = "Failed";
						
					}
					

				} catch (FileNotFoundException e) {
					e.printStackTrace();
				} catch (IOException e) {
					e.printStackTrace();
				}
						
				return FinalStatus;
			}
			
			//Get the HTML reporter executionTime
			public static String fn_Get_Test_ExecutionTime(String HTMLFilePath) throws Exception {
				
				String content = null;
				String SplitContent = null;
				String finalExecutionTime = null;
				
				try {
					
					FileInputStream fileinput = new FileInputStream(HTMLFilePath);
					content = IOUtils.toString(fileinput, "UTF-8");
					SplitContent = ">EXECUTION DURATION </td><td>";
					String[] SplitValue = content.split(SplitContent);
					
					//Find the first occurrence of the string '<'
					int position = SplitValue[1].indexOf("<", 0);
					
					try {
						
						//Get the Fail Count
						String ExecutionTime = SplitValue[1].substring(0, position);
						System.out.println(ExecutionTime);
						finalExecutionTime = ExecutionTime.replace("00 hr","12").replace(" min", "").replace("sec", "AM");
						
					} catch (Exception e) {
						e.printStackTrace();
					}
					

				} catch (FileNotFoundException e) {
					e.printStackTrace();
				} catch (IOException e) {
					e.printStackTrace();
				}
						
				return finalExecutionTime;
			}
			
			//Get the HTML reporter execution start time
			public static String fn_Get_Test_ExecutionStartTime(String HTMLFilePath) throws Exception {
				
				String content = null;
				String SplitContent = null;
				String ExecutionStartTime = null;
				
				try {
					
					FileInputStream fileinput = new FileInputStream(HTMLFilePath);
					content = IOUtils.toString(fileinput, "UTF-8");
					SplitContent = ">EXECUTION STARTED ON </td><td>";
					String[] SplitValue = content.split(SplitContent);
					
					//Find the first occurrence of the string '<'
					int position = SplitValue[1].indexOf("<", 0);
					
					try {
						
						//Get the Fail Count
						ExecutionStartTime = SplitValue[1].substring(0, position);
						
					} catch (Exception e) {
						e.printStackTrace();
					}
					

				} catch (FileNotFoundException e) {
					e.printStackTrace();
				} catch (IOException e) {
					e.printStackTrace();
				}
						
				return ExecutionStartTime;
			}
			
			//Get the HTML reporter execution start time
			public static int fn_Get_Test_ExecutionPassCount(String HTMLFilePath) throws Exception {
				
				String content = null;
				String SplitContent = null;
				int strPassCount = 0;
				
				try {
					
					FileInputStream fileinput = new FileInputStream(HTMLFilePath);
					content = IOUtils.toString(fileinput, "UTF-8");
					SplitContent = ">PASS </font></td><td class=\"2\">";
					String[] SplitValue = content.split(SplitContent);
					
					//Find the first occurrence of the string '<'
					int position = SplitValue[1].indexOf("<", 0);
					
					try {
						
						//Get the Fail Count
						strPassCount = Integer.valueOf(SplitValue[1].substring(0, position));
						
					} catch (Exception e) {
						e.printStackTrace();
					}
					

				} catch (FileNotFoundException e) {
					e.printStackTrace();
				} catch (IOException e) {
					e.printStackTrace();
				}
						
				return strPassCount;
			}
			
			//Get the HTML reporter execution start time
			public static int fn_Get_Test_ExecutionFailCount(String HTMLFilePath) throws Exception {
				
				String content = null;
				String SplitContent = null;
				int strFailCount = 0;
				
				try {
					
					FileInputStream fileinput = new FileInputStream(HTMLFilePath);
					content = IOUtils.toString(fileinput, "UTF-8");
					SplitContent = ">FAIL </font></td><td class=\"2\">";
					String[] SplitValue = content.split(SplitContent);
					
					//Find the first occurrence of the string '<'
					int position = SplitValue[1].indexOf("<", 0);
					
					try {
						
						//Get the Fail Count
						strFailCount = Integer.valueOf(SplitValue[1].substring(0, position));
						
					} catch (Exception e) {
						e.printStackTrace();
					}
					

				} catch (FileNotFoundException e) {
					e.printStackTrace();
				} catch (IOException e) {
					e.printStackTrace();
				}
						
				return strFailCount;
			}
			
			//Update the Excel Report
			public static void fn_Update_Reporter(ArrayList<String> Ticker, ArrayList<String> Component, ArrayList<String> Status, ArrayList<String> HTMLPath) throws SQLException {
				
				try {
					
					ArrayList<String> Failed_Fund = new ArrayList<String>();
					ArrayList<String> Failed_Script = new ArrayList<String>();
					
					String ORADEVPIM_URL1 = null;
					String UID1 = null;
					String PWD1 = null;
					String InsertFailedSQLQuery = null;
					
					//System.out.println("Updating Component Level Status");
					
					FileInputStream Fileip = new FileInputStream(AppConstant.BatchReporterExcel);
					HSSFWorkbook workbookip = new HSSFWorkbook(Fileip);
					HSSFSheet Sheetip = workbookip.getSheet("Dashboard");
					
					CreationHelper createHelper = workbookip
					.getCreationHelper();
											
					//Create the Row for the Header
					HSSFRow HeaderRow = Sheetip.createRow(6);
					
					//Create the Cell for the Header
					HSSFCell HeadCell1 = HeaderRow.createCell(0);
					HSSFCell HeadCell2 = HeaderRow.createCell(1);
					HSSFCell HeadCell3 = HeaderRow.createCell(2);
					HSSFCell HeadCell4 = HeaderRow.createCell(3);
					
					//Input the Headers
					HeadCell1.setCellValue("Sr. #");
					HeadCell2.setCellValue("ETF Ticker");
					HeadCell3.setCellValue("Component Tab");
					HeadCell4.setCellValue("Status");
					
					//Get the cell for the super Header
					HSSFRow SuperHeaderRow = Sheetip.getRow(0);
					HSSFCell SuperHeaderCell = SuperHeaderRow.getCell(0);
					
					//Set the super header
					SuperHeaderCell.setCellValue("ETF Websites – QA Automated Test Results");
					
					//Create the font for the super header
					Font SuperHeaderFont = workbookip.createFont();
					CellStyle SuperHeaderStyle = workbookip.createCellStyle();
					
					//Set the super Header Font to 'calibri'
					SuperHeaderFont.setFontName("calibri");
					SuperHeaderFont.setFontHeightInPoints((short) (16));
					
					//Set the super header cell style
					SuperHeaderStyle.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
					SuperHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); //CellStyle.SOLID_FOREGROUND);
					SuperHeaderStyle.setBorderBottom(BorderStyle.THIN); //CellStyle.BORDER_THIN);
					SuperHeaderStyle.setBorderTop(BorderStyle.THIN); //CellStyle.BORDER_THIN);
					SuperHeaderStyle.setBorderLeft(BorderStyle.THIN); //CellStyle.BORDER_THIN);
					SuperHeaderStyle.setBorderRight(BorderStyle.THIN); //CellStyle.BORDER_THIN);
					SuperHeaderStyle.setAlignment(HorizontalAlignment.CENTER); //CellStyle.ALIGN_CENTER);
					SuperHeaderStyle.setFont(SuperHeaderFont);
					SuperHeaderFont.setColor(IndexedColors.WHITE.getIndex());
					
					SuperHeaderCell.setCellStyle(SuperHeaderStyle);
					
					//Create the font for the table header
					Font HeaderFont = workbookip.createFont();
					CellStyle HeaderStyle = workbookip.createCellStyle();
					
					//Set the Header Font to 'calibri'
					HeaderFont.setFontName("calibri");
					HeaderFont.setFontHeightInPoints((short) (11));
								
					HeaderStyle.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
					HeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); //CellStyle.SOLID_FOREGROUND);
					HeaderStyle.setBorderBottom(BorderStyle.THIN); //CellStyle.BORDER_THIN);
					HeaderStyle.setBorderTop(BorderStyle.THIN); //CellStyle.BORDER_THIN);
					HeaderStyle.setBorderLeft(BorderStyle.THIN); //CellStyle.BORDER_THIN);
					HeaderStyle.setBorderRight(BorderStyle.THIN); //CellStyle.BORDER_THIN);
					HeaderStyle.setAlignment(HorizontalAlignment.CENTER); //CellStyle.ALIGN_CENTER);
					HeaderStyle.setFont(HeaderFont);
					HeaderFont.setColor(IndexedColors.WHITE.getIndex());
					
					HeadCell1.setCellStyle(HeaderStyle);
					HeadCell2.setCellStyle(HeaderStyle);
					HeadCell3.setCellStyle(HeaderStyle);
					HeadCell4.setCellStyle(HeaderStyle);
					
					int columnIndex1 = HeadCell1.getColumnIndex();
					int columnIndex2 = HeadCell2.getColumnIndex();
					int columnIndex3 = HeadCell3.getColumnIndex();
					int columnIndex4 = HeadCell4.getColumnIndex();
					
					Sheetip.autoSizeColumn(columnIndex1);
					Sheetip.autoSizeColumn(columnIndex2);
					Sheetip.autoSizeColumn(columnIndex3);
					Sheetip.autoSizeColumn(columnIndex4);
					
					//Create the Font
					//Font Component_Serial_Number_Font = workbookip.createFont();
					//Font Component_ETF_Ticker_Font = workbookip.createFont();
					//Font Component_Component_Font = workbookip.createFont();
					Font Component_Pass_Status_Font = workbookip.createFont();
					Font Component_Fail_Status_Font = workbookip.createFont();
					Font Normal_Font = workbookip.createFont();
					
					//Create Cell Style
					//CellStyle Component_Serial_Number = workbookip.createCellStyle();
					//CellStyle Component_ETF_Ticker = workbookip.createCellStyle();
					//CellStyle Component_Component = workbookip.createCellStyle();
					CellStyle Component_Pass_Status = workbookip.createCellStyle();
					CellStyle Component_Fail_Status = workbookip.createCellStyle();
					CellStyle Normal_CellStyle = workbookip.createCellStyle();
					
					//Set the normal font
					Normal_Font.setFontName("calibri");
					Normal_Font.setColor(IndexedColors.BLACK.getIndex());
					Normal_CellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
					Normal_CellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); //CellStyle.SOLID_FOREGROUND);
					Normal_CellStyle.setBorderBottom(BorderStyle.THIN); //CellStyle.BORDER_THIN);
					Normal_CellStyle.setBorderTop(BorderStyle.THIN); //CellStyle.BORDER_THIN);
					Normal_CellStyle.setBorderLeft(BorderStyle.THIN); //CellStyle.BORDER_THIN);
					Normal_CellStyle.setBorderRight(BorderStyle.THIN); //CellStyle.BORDER_THIN);
					Normal_CellStyle.setAlignment(HorizontalAlignment.CENTER); //CellStyle.ALIGN_CENTER);
					Normal_CellStyle.setFont(Normal_Font);
								
					HSSFPalette passpalette = workbookip.getCustomPalette();
					HSSFColor PassColor = passpalette.findSimilarColor(51, 255, 51);
					short PassPalindex = PassColor.getIndex();
					
					HSSFPalette failpalette = workbookip.getCustomPalette();
					HSSFColor FailColor = failpalette.findSimilarColor(255, 51, 51);
					short FailPalindex = FailColor.getIndex();
					
					Component_Pass_Status_Font.setFontName("calibri");
					Component_Pass_Status_Font.setUnderline(HSSFFont.U_SINGLE);
					//Component_Pass_Status.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
					Component_Pass_Status.setFillForegroundColor(PassPalindex);
					Component_Pass_Status.setFillPattern(FillPatternType.SOLID_FOREGROUND); //CellStyle.SOLID_FOREGROUND);
					Component_Pass_Status.setBorderBottom(BorderStyle.THIN); //CellStyle.BORDER_THIN);
					Component_Pass_Status.setBorderTop(BorderStyle.THIN); //CellStyle.BORDER_THIN);
					Component_Pass_Status.setBorderLeft(BorderStyle.THIN); //CellStyle.BORDER_THIN);
					Component_Pass_Status.setBorderRight(BorderStyle.THIN); //CellStyle.BORDER_THIN);
					Component_Pass_Status.setAlignment(HorizontalAlignment.CENTER); //CellStyle.ALIGN_CENTER);
					Component_Pass_Status.setFont(Component_Pass_Status_Font);
					
					Component_Fail_Status_Font.setFontName("calibri");
					Component_Fail_Status_Font.setUnderline(HSSFFont.U_SINGLE);
					//Component_Fail_Status.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
					Component_Fail_Status.setFillForegroundColor(FailPalindex);
					Component_Fail_Status.setFillPattern(FillPatternType.SOLID_FOREGROUND); //CellStyle.SOLID_FOREGROUND);
					Component_Fail_Status.setBorderBottom(BorderStyle.THIN); //CellStyle.BORDER_THIN);
					Component_Fail_Status.setBorderTop(BorderStyle.THIN); //CellStyle.BORDER_THIN);
					Component_Fail_Status.setBorderLeft(BorderStyle.THIN); //CellStyle.BORDER_THIN);
					Component_Fail_Status.setBorderRight(BorderStyle.THIN); //CellStyle.BORDER_THIN);
					Component_Fail_Status.setAlignment(HorizontalAlignment.CENTER); //CellStyle.ALIGN_CENTER);
					Component_Fail_Status.setFont(Component_Fail_Status_Font);
					
					int RowIndex = 7;
					int columnIndex = 0;
					
					//Clear the Array
					Failed_Fund.clear();
					Failed_Script.clear();
					
					for (int tkr = 0; tkr < Ticker.size(); tkr++) {
						
						HSSFRow Row = Sheetip.createRow(RowIndex);
						
						columnIndex = 0;
						
						for (int Totcol = 0; Totcol < 4; Totcol++) {		
							
							columnIndex = 0;					
							
							//Input the values
							switch (Totcol) {
							
							case 0:
								
								HSSFCell SerialNoCell = Row.createCell(Totcol);
								
								//Convert the int to String
								String SNo = String.valueOf(tkr+1);
								SerialNoCell.setCellValue(SNo);
								
								//Set the Style
								SerialNoCell.setCellStyle(Normal_CellStyle);
								
								//Get the Column Index and Align the text
								columnIndex = SerialNoCell.getColumnIndex();
								Sheetip.autoSizeColumn(columnIndex);
								
								break;

							case 1:
								
								HSSFCell TickerCell = Row.createCell(Totcol);
								
								TickerCell.setCellValue(Ticker.get(tkr));
								
								//Set the Style
								TickerCell.setCellStyle(Normal_CellStyle);
								
								//Get the Column Index and Align the text
								columnIndex = TickerCell.getColumnIndex();
								Sheetip.autoSizeColumn(columnIndex);
								
								break;
								
							case 2:
								
								HSSFCell ComponentCell = Row.createCell(Totcol);
								
								ComponentCell.setCellValue(Component.get(tkr));
								
								//Set the Style
								ComponentCell.setCellStyle(Normal_CellStyle);
								
								//Get the Column Index and Align the text
								columnIndex = ComponentCell.getColumnIndex();
								Sheetip.autoSizeColumn(columnIndex);
								
								break;
								
							case 3:
								
								HSSFCell StatusCell = Row.createCell(Totcol);
								StatusCell.setCellValue(Status.get(tkr));
								
								HSSFHyperlink HTML_FILE_Link = (HSSFHyperlink) createHelper.createHyperlink(HyperlinkType.FILE); //.createHyperlink(HSSFHyperlink.LINK_FILE);
								HTML_FILE_Link.setAddress(HTMLPath.get(tkr)); 
								StatusCell.setHyperlink(HTML_FILE_Link);
																				
								//Set the cell color according to Pass/Fail
								if (Status.get(tkr).contains("Pass")) {
									StatusCell.setCellStyle(Component_Pass_Status);
								} else {
									StatusCell.setCellStyle(Component_Fail_Status);
									Failed_Fund.add(Ticker.get(tkr));
									Failed_Script.add(Component.get(tkr));
								}
								
								//Get the Column Index and Align the text
								columnIndex = StatusCell.getColumnIndex();
								Sheetip.autoSizeColumn(columnIndex);
								
								break;
								
							default:
								break;
								
							}
							
						}
						
						RowIndex ++;
						
					}
															
					//System.out.println("Completed -> Updating Component Level Status");
					
				} catch (IOException e) {
					e.printStackTrace();
				}
				
				//System.out.println("Completed Updating Failed DataPoint Entry");
						
			}

			//Update the Excel Report
			public static void fn_Update_Reporter1(ArrayList<String> ScriptName, ArrayList<String> Status, ArrayList<String> HTMLPath, ArrayList<String> AutomationExecTime, String AutomationStartTime, int TotalPass, int TotalFail) throws SQLException {
				
				
				try {
					
					FileInputStream Fileip = new FileInputStream(AppConstant.BatchReporterExcel);
					HSSFWorkbook workbookip = new HSSFWorkbook(Fileip);
					HSSFSheet Sheetip = workbookip.getSheet("Dashboard");
					CreationHelper createHelper = workbookip.getCreationHelper();
					
					HSSFCellStyle cellStatusStyle = workbookip.createCellStyle();
					Font statusFont = workbookip.createFont();
					statusFont.setUnderline(HSSFFont.U_SINGLE);
					statusFont.setFontHeightInPoints((short)9);
					cellStatusStyle.setFont(statusFont);
					cellStatusStyle.setAlignment(HorizontalAlignment.CENTER); //CellStyle.ALIGN_CENTER);
					HSSFRow Row = Sheetip.getRow(3);
					Iterator<Row> itrRow = Sheetip.rowIterator();
					
					//Set the ExecutionStartTime
					HSSFCell cell = Row.getCell(3);
					cell = Row.getCell(2);
					cell.setCellValue(cell.getStringCellValue()+" "+AutomationStartTime);
										
					Row = Sheetip.getRow(29);
					for (int itr = 0; itr < ScriptName.size(); itr++) {	
						
					
						//Set the serial number
						cell = Row.getCell(0);		
						cell.setCellValue(itr+1);
						
						//Set the scenario name
						cell = Row.getCell(1);
						cell.setCellValue(ScriptName.get(itr));
						
						//Set the status
						cell = Row.getCell(9);
						cell.setCellValue(Status.get(itr));
						cell.setCellStyle(cellStatusStyle);
						HSSFHyperlink HTML_FILE_Link = (HSSFHyperlink) createHelper.createHyperlink(HyperlinkType.FILE); //.createHyperlink(HSSFHyperlink.LINK_FILE);
						HTML_FILE_Link.setAddress(HTMLPath.get(itr)); 
						cell.setHyperlink(HTML_FILE_Link);
						
						//Set the execution time
						cell = Row.getCell(12);
						cell.setCellValue(AutomationExecTime.get(itr));
						
						//Point to the next row
						Row = (HSSFRow) itrRow.next();
					}
					
					//Set the ExecutionStartTime
					Row = Sheetip.getRow(9);
					cell = Row.getCell(2);
					cell.setCellValue(TotalPass);
					
					cell = Row.getCell(3);
					cell.setCellValue(TotalFail);
					
					//Save the output
					FileOutputStream fop = new FileOutputStream(AppConstant.BatchReporterExcel);
					workbookip.write(fop);
					workbookip.close();
				}catch (IOException e) {
					e.printStackTrace();
				}
			}
}
