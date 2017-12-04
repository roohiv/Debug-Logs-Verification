package Scanr;

import java.awt.AWTException;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Set;

//import javax.imageio.ImageIO;















import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Action;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.internal.WrapsDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class AutomateMain {

	static public String workFlowLink;
	static public String login_id;
	static public String login_password;
	static public String flow_Name;
	static public String SO_Name;
	static public String ban_No_Parameters;
	static public String RR_Name;
	static public String Display_Name;
	static public String SO_Result;
	static public String outReportingPath;
	static public String Interface_Api_Name;
	static public String ParameterName_Value;
 	static public String Interface_Api_Name_Status;
	static public String outputPath;
	static boolean browserExist=false;
	static public int imageNameCounter=0;
	static public String image_Name;
	static public String startTime;
	static public String endTime;
	static public String testName;
	static public XSSFCell reportingEndTime;

	public String status="";
	private static WebDriver driver;
	private static WebDriverWait wait;
	
	public void Read(String ExcelPath) throws IOException, InterruptedException {
		File excel = new File(ExcelPath);
		outReportingPath  = ExcelPath.substring(0,ExcelPath.lastIndexOf('\\')+1);
		outputPath=outReportingPath+ExcelPath.substring(ExcelPath.lastIndexOf('\\')+1,ExcelPath.lastIndexOf('.'));
		FileInputStream fin = new FileInputStream(excel);
		XSSFWorkbook wb = new XSSFWorkbook(fin);
		XSSFSheet ws = wb.getSheetAt(0);

		// Create the spreadsheet    
		XSSFSheet reportingSheet = wb.getSheetAt(1);

		/// CREATING HEADER REPORT SHEET ////////////////
		createHeaderReport(reportingSheet);

		
		XSSFName n1 = wb.getName("Step_Action");
		XSSFName n2 = wb.getName("Test_Name");
		XSSFName n3 = wb.getName("Execution");
		XSSFName n4 = wb.getName("Merged_Test_Cases");
		XSSFName n5 = wb.getName("Track");
		CellReference ref1=  new CellReference(n1.getRefersToFormula());
		CellReference ref2=  new CellReference(n2.getRefersToFormula());
		CellReference ref3=  new CellReference(n3.getRefersToFormula());
		CellReference ref4=  new CellReference(n4.getRefersToFormula());
		CellReference ref5=  new CellReference(n5.getRefersToFormula());
		int rowNum = ref1.getRow()+1;
		int colNum = ref1.getCol();
		int colNumTestName = ref2.getCol();
		int colExecuteTC = ref3.getCol();
		int colMergedTestCase=ref4.getCol();
		int colTrack=ref5.getCol();
		int rowNumLen = ws.getLastRowNum() + 1;
		XSSFCell cell,cellTestName,reportingTestName,reportingStatus,cellIsExecute,reportingReason,reportingStartTime,mergeTest,track,reportMergeTest,reportTrack,reportingScreenshotLocation;
		XSSFRow row2,row;
		for (int i = rowNum; i < rowNumLen; i++) {
			admin_Logs auto = new admin_Logs();
			row = ws.getRow(i);
			/////////////     for converting to line 2 , i-1 -> i     //////////////			
			row2 = reportingSheet.createRow(i);			
			cell = row.getCell(colNum);
			cellTestName = row.getCell(colNumTestName);
			cellIsExecute = row.getCell(colExecuteTC);
			mergeTest=row.getCell(colMergedTestCase);
			track=row.getCell(colTrack);		
			
			reportMergeTest=row2.createCell(0);
			reportTrack=row2.createCell(1);
			reportingTestName = row2.createCell(2);
			reportingStatus = row2.createCell(3);
			reportingReason = row2.createCell(4);
			reportingStartTime=row2.createCell(5);
			reportingEndTime=row2.createCell(6);
			reportingScreenshotLocation=row2.createCell(7);
			// XSSFCell cell2 = row.getCell(j + 1);

			if (cell !=null && cellIsExecute!=null && cellIsExecute.toString().equalsIgnoreCase("YES"))
			{
/////////////////////////////////   MAKING DIR //////////////////////////////////
				testName=cellTestName.toString();
//				new File(outReportingPath+testName).mkdirs();
				File theDir = new File(outputPath);
				if (!theDir.exists()) {
					theDir.mkdir();     
				}
				File screenDir = new File(outputPath+"\\"+testName);
				if (!screenDir.exists()) {
					screenDir.mkdir();     
				}
				image_Name = outputPath+"\\"+testName+"\\"+imageNameCounter+".png";
				
				
				//reportMergeTest.setCellValue(mergeTest.toString());
				//reportTrack.setCellValue(track.toString());
				//reportingTestName.setCellValue(cellTestName.toString());
				//reportingScreenshotLocation.setCellValue(outputPath+"\\"+testName);
				String[] array = cell.toString().split("\\n");
				ArrayList<String> step_Listed_list = new ArrayList<String>();
				for(String s: array)
				{
					s=s.trim();
					if (s.length()>1)
					{
						if (Character.isDigit(s.charAt(0)))
						{
							step_Listed_list.add(s);
						}
					}

					// System.out.println(cell2.toString());
					// if (cell2.toString().equals("Y")) {
				}
				String[] stockArr = new String[step_Listed_list.size()];
				stockArr = step_Listed_list.toArray(stockArr);

				for(String as : stockArr)
				{
/////////////////////////////////      COMMON      ///////////////////////////////////////////////

					if (as	.toLowerCase().contains("Open".toLowerCase()) && as	.toLowerCase().contains("WFB".toLowerCase()) && as.toLowerCase().contains("url".toLowerCase())&& driver==null)
					{
						//						startTime=getTime();	
						//						reportingStartTime.setCellValue(startTime);
						workFlowLink = as.substring(as.indexOf(':')+1,as.lastIndexOf('.')).trim();
						OpenWFBUrl(workFlowLink);

					}
					else if (as.toLowerCase().contains("Login".toLowerCase()) && as.toLowerCase().contains("with".toLowerCase()) && as.toLowerCase().contains("User".toLowerCase()) && as.toLowerCase().contains("Role".toLowerCase())&&!browserExist)
					{
						browserExist=true;
						String temp = as.substring(as.indexOf(':')+1,as.lastIndexOf('.')).trim();
						String tempArray [] = temp.split("/");
						login_id = tempArray[0];
						login_password = tempArray[1];
						Login(login_id, login_password);
					}


/////////////////////////////////     CHECK RR   ////////////////////////////////////////////////				

					else if (as.toLowerCase().contains("Search".toLowerCase())&& as.toLowerCase().contains("open".toLowerCase()) && as.toLowerCase().contains("Workflow".toLowerCase()))
					{
						flow_Name = as.substring(as.indexOf(':')+1,as.lastIndexOf('.')).trim();
						flow_Name=flow_Name.replace(".", "").trim();
						System.out.println("START TIME : ");
						startTime=getTime();							
						reportingStartTime.setCellValue(startTime);
						SearchWorkFlow(flow_Name);
						OpenWorkFlow(flow_Name);
						Thread.sleep(1500);
						OpenAutoSave(true);
						ChangeTab();
					}
					
					else if (as.toLowerCase().contains("Input".toLowerCase())&& as.toLowerCase().contains("Values".toLowerCase()))
					{
						wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(@class,\"x-btn-text\")and contains(@title,'Opens a panel to test workflows with sample inputs.')]")));
						int count1;
						int count2;
						int m;
						int n;
     					int p=0;
     					int MOdelSize=0;
     					int MOdelSize1=0;
     					int Counter=0;
						String val = as.substring(as.indexOf(':')+1,as.lastIndexOf('.')).trim();
						String[] val11 = val.split("\\;");
						for(String s2:val11){
							String[] One=s2.split("\\&");
							String DeviceName=One[0].trim();
							String DeviceInput=One[1].trim();
							String[] val1 = One[1].split("\\@");
						wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("svg")));
						WebElement SVG=driver.findElement(By.cssSelector("svg"));
						List<WebElement> li=SVG.findElements(By.cssSelector("tspan"));
						int count=li.size();
						System.out.println(count);
						for(WebElement w : li){
							String str=w.getText().trim();
							if(DeviceName.startsWith(str)){
								w.click();
								List<WebElement> Tabb=driver.findElements(By.cssSelector("table[class='x-grid3-row-table']"));
								count1=Tabb.size();
								System.out.println(count1);
								for(m=0;m<=count1-1;m++){
								Tabb=driver.findElements(By.cssSelector("table[class='x-grid3-row-table']"));
								count1=Tabb.size();
								List<WebElement> im=Tabb.get(m).findElements(By.cssSelector("div"));
								count2=im.size();
								System.out.println(count2);
								for(n=0;n<=count2-3;n++){
									Tabb=driver.findElements(By.cssSelector("table[class='x-grid3-row-table']"));
									Actions act=new Actions(driver);
									String str1=Tabb.get(m).findElements(By.cssSelector("div")).get(n).getText().trim();
									System.out.println(str1);
									if(driver.findElements(By.cssSelector("div[class$='x-window-draggable']")).size()!=0){
											int size=driver.findElements(By.cssSelector("div[class$='x-window-draggable']")).size();
											for(int s=0;s<=size-1;s++){
												String ClassName=driver.findElements(By.cssSelector("div[class$='x-window-draggable']")).get(s).findElement(By.cssSelector("span[class='x-window-header-text']")).getText();
												if(ClassName.matches("Error")){
												
											       driver.findElements(By.cssSelector("div[class$='x-window-draggable']")).get(s).findElement(By.cssSelector("div[class$='x-tool-close']")).click();
												}
												}
												
												
										
										
									}
									String Value1233=Tabb.get(m).findElements(By.cssSelector("div")).get(n).getText();
									act.doubleClick(Tabb.get(m).findElements(By.cssSelector("div")).get(n)).build().perform();
									Thread.sleep(1000);
									if(driver.findElements(By.cssSelector("img[class$='icon-modelinput']")).size()!=0){
										MOdelSize=driver.findElements(By.cssSelector("img[class$='icon-modelinput']")).size();
										p=p+1;
										
									}
									
									
									
									for(String s:val1){
										String[] val2 = s.split("\\#");
										String InputField=val2[0].trim();
										String InputFieldValue=val2[1].trim();
										if (str1.equalsIgnoreCase(InputField)){
											
											n=n+1;	
											Value1233=Tabb.get(m).findElements(By.cssSelector("div")).get(n).getText();
											act.doubleClick(Tabb.get(m).findElements(By.cssSelector("div")).get(n)).build().perform();
											Thread.sleep(1000);
											if(InputField.equalsIgnoreCase("Action Mode")){
											if(driver.findElements(By.cssSelector("img[class$='x-form-arrow-trigger']")).size()!=0){
											  MOdelSize1=driver.findElements(By.cssSelector("img[class$='x-form-arrow-trigger']")).size();
											  driver.findElements(By.cssSelector("img[class$='x-form-arrow-trigger']")).get(MOdelSize1-1).click();
											  List<WebElement> ltr=driver.findElements(By.cssSelector("div[class^='x-combo-list-item']"));
											 int Comsize=driver.findElements(By.cssSelector("div[class^='x-combo-list-item']")).size();
											 // int Comsize=t1.findElements(By.cssSelector("div[class$='x-combo-list']")).size();
											 for(int e1=0;e1<=Comsize-1;e1++){
											 // int ElSize=findElements(By.xpath("div[class='x-combo-list-inner']//div")).get(e1).findElements(By.cssSelector("div")).size();
											  //for(int e11=0;e11<=ElSize-1;e11++){
											    String Test=ltr.get(e1).getText().trim();
											  if(Test.matches(InputFieldValue)){
												  ltr.get(e1).click();
											  }
											 // }
											  }
											  }
											}
											  
											if(driver.findElements(By.cssSelector("img[class$='icon-modelinput']")).size()!=0){
												if(driver.findElements(By.cssSelector("div[class$='x-window-draggable']")).size()!=0){
													int size=driver.findElements(By.cssSelector("div[class$='x-window-draggable']")).size();
													for(int s1=0;s1<=size-1;s1++){
														String ClassName=driver.findElements(By.cssSelector("div[class$='x-window-draggable']")).get(s1).findElement(By.cssSelector("span[class='x-window-header-text']")).getText();
														if(ClassName.matches("Error")){
														
													       driver.findElements(By.cssSelector("div[class$='x-window-draggable']")).get(s1).findElement(By.cssSelector("div[class$='x-tool-close']")).click();
														}
														}
														
												MOdelSize=driver.findElements(By.cssSelector("img[class$='icon-modelinput']")).size();
												driver.findElements(By.cssSelector("img[class$='icon-modelinput']")).get(MOdelSize-1).click();
												ChangeTab();
												WebElement str11=driver.findElement(By.cssSelector("table[class='mceLayout']"));
												WebElement w3=str11.findElement(By.cssSelector("iframe"));
													driver.switchTo().frame(w3);
													 wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("body")));
													 WebElement element=driver.findElement(By.xpath("//body[@id='tinymce']"));
													 element.clear();
													 element.sendKeys(InputFieldValue);
													 driver.switchTo().defaultContent();
													 WebElement Tabb1=driver.findElement(By.cssSelector("table[class='x-btn-wrap x-btn']"));
													 List<WebElement> dw =Tabb1.findElements(By.cssSelector("button[class='x-btn-text']"));
														int count3=dw.size();
														Thread.sleep(1000);
														System.out.println(count3);
													 for(int r=0;r<=count3-1;r++){
														 Thread.sleep(1000);
														 String S112=driver.findElement(By.cssSelector("table[class='x-btn-wrap x-btn']")).findElements(By.cssSelector("button[class='x-btn-text']")).get(r).getText().trim();
														 System.out.println(S112);
														 if (S112.matches("Ok")){
															 driver.findElement(By.cssSelector("table[class='x-btn-wrap x-btn']")).findElements(By.cssSelector("button[class='x-btn-text']")).get(r).click();
															 Thread.sleep(1000);
															 break;
															 
														 }
														 ErrorDescripterHandling1();
											}
													
													 ErrorDescripterHandling1();		
											}
												ErrorDescripterHandling1();
											}
											
											else{
												
												
											}
											
											ErrorDescripterHandling1();
									
											 
									}
										ErrorDescripterHandling1();
								}
									ErrorDescripterHandling1();
								}
								ErrorDescripterHandling1();
							}
								ErrorDescripterHandling1();
							}
						

								
															
									
									
									else{
										
									}
							ErrorDescripterHandling1();
									}
						}
						
						Thread.sleep(10000);
						ErrorDescripterHandling1();
						SaveInputs(true);
						Thread.sleep(2000);
					}
									

							
						
					
						
					

					

					else if (as.toLowerCase().contains("enter".toLowerCase()) && as.toLowerCase().contains("BAN".toLowerCase())&&as.toLowerCase().contains("Launch".toLowerCase()) && as.toLowerCase().contains("workflow".toLowerCase()))
					{
						
						ErrorDescripterHandling1();
						ban_No_Parameters = as.substring(as.indexOf(':')+1,as.lastIndexOf('.')).trim();
						Thread.sleep(2000);
						
						
						
						
						OutputStream log=auto.runCommands(ExcelPath,outputPath+"\\"+testName);
						
						
						//				BreakLock(true);
						LaunchFlow(flow_Name,ban_No_Parameters);
						auto.Disconnect(log);
						continueClickTillSuccess();
						//						if (flow_Name.equalsIgnoreCase("A_UFO_GETNRCharge"))
						//						{
						//							Run_A_UFO_GETNRCharge();
						//						}
						//						else if (flow_Name.equalsIgnoreCase("Hitesh_WAN"))
						//						{
						//							Run_WANTest();
						//						}
						//						else if (flow_Name.equalsIgnoreCase("UtilLSCreateAOTSTicket_Sashank"))
						//						{
						//							Run_UtilLSCreateAOTSTicket_Sashank();
						//						}
						
					}
					else if(as.toLowerCase().contains("Verify".toLowerCase()) && as .toLowerCase().contains("Request".toLowerCase()))
					{
						Interface_Api_Name_Status = as.substring(as.indexOf(':')+1,as.lastIndexOf('@')).trim();
						String text=auto.runCommands_s2(testName,i,ban_No_Parameters,Interface_Api_Name_Status);
						Interface_Api_Name = as.substring(as.indexOf(':')+1,as.lastIndexOf('!')).trim();
						ParameterName_Value = as.substring(as.indexOf('@')+1,as.lastIndexOf('.')).trim();
						auto.Verify_Request_Tags(text,i,ParameterName_Value,Interface_Api_Name);
					}
					
					else if (as.toLowerCase().contains("open".toLowerCase()) && as .toLowerCase().contains("dictionary".toLowerCase()))
					{
						SO_Name = as.substring(as.indexOf(':')+1,as.lastIndexOf('.')).trim();
						OpenDictionary();
					}
					
					else if (as.toLowerCase().contains("Flow".toLowerCase()) && as .toLowerCase().contains("Ends".toLowerCase()))
					{
						flowEnds();
					}
					

					else if (as.toLowerCase().contains("checkrr"))
					{
						RR_Name = as.substring(as.indexOf(':')+1,as.lastIndexOf('.')).trim();
						String result=CheckRR(SO_Name,RR_Name);
						if(result.contains("Success"))
						{
							reportingStatus.setCellValue("Pass");
						}
						else
						{
							reportingStatus.setCellValue("Fail");
						}
						reportingReason.setCellValue(result);
						Thread.sleep(2000);
						flowEnds();
					}
					else if (as.toLowerCase().contains("checkdisplayname"))
					{
						Display_Name=as.substring(as.indexOf(':')+1,as.lastIndexOf('.')).trim();						
						String result=checkDisplayName(SO_Name,Display_Name.toLowerCase());
						if(result.contains("Success"))
						{
							reportingStatus.setCellValue("Pass");
						}
						else
						{
							reportingStatus.setCellValue("Fail");
						}
						reportingReason.setCellValue(result);
						flowEnds();
					}
					else if (as.toLowerCase().contains("checkresult"))
					{
						SO_Result=as.substring(as.indexOf(':')+1,as.lastIndexOf('.')).trim();
						String result=checkResult(SO_Name, SO_Result);
						if(result.contains("Success"))
						{
							reportingStatus.setCellValue("Pass");
						}
						else
						{
							reportingStatus.setCellValue("Fail");
						}
						reportingReason.setCellValue(result);
						flowEnds();

					}
					else if (as.toLowerCase().contains("checkerror"))
					{
						String result=checkError(SO_Name);
						if(result.contains("Success"))
						{
							reportingStatus.setCellValue("Pass");
						}
						else
						{
							reportingStatus.setCellValue("Fail");
						}
						reportingReason.setCellValue(result);
						flowEnds();
					}
					else if (as.toLowerCase().contains("checkproperties"))
					{
						String properties=as.substring(as.indexOf(':')+1,as.lastIndexOf('.')).trim();
						String[] propArray=properties.split("#");
						//						System.out.println(propArray[0]+","+propArray[1]);
						HashMap<String, String> propMap= new HashMap<String, String>();
						for(String property:propArray)
						{
							String[] keyVal=property.split(",");
							propMap.put(keyVal[0], keyVal[1]);	 
						}
						String result=checkProperties(SO_Name, propMap);
						if(result.toLowerCase().contains("fail"))
						{
							reportingStatus.setCellValue("Fail");
						}
						else
						{
							reportingStatus.setCellValue("Pass");
						}
						reportingReason.setCellValue(result);
						flowEnds();
					}
					else if (as.toLowerCase().contains("checkmultiplepropertiesinitLS"))
					{
						String properties=as.substring(as.indexOf(':')+1,as.lastIndexOf('.')).trim();
						//String result=checkMultipleProperties(SO_Name,properties);
						String result=checkMultiplePropertiesInitLS(SO_Name,properties);
						if(result.toLowerCase().contains("fail"))
						{
							reportingStatus.setCellValue("Fail");
						}
						else
						{
							reportingStatus.setCellValue("Pass");
						}
						reportingReason.setCellValue(result);
						Thread.sleep(2000);
						flowEnds();
					}
					
					else if (as.toLowerCase().contains("checkmultipleproperties"))
					{
						String properties=as.substring(as.indexOf(':')+1,as.lastIndexOf('.')).trim();
						//String result=checkMultipleProperties(SO_Name,properties);
						String result=checkMultipleProperties(SO_Name,properties);
						if(result.toLowerCase().contains("fail"))
						{
							reportingStatus.setCellValue("Fail");
						}
						else
						{
							reportingStatus.setCellValue("Pass");
						}
						reportingReason.setCellValue(result);
						Thread.sleep(2000);
						flowEnds();
					}
				

					/*					else if (as.toLowerCase().contains("enter".toLowerCase()) && as.toLowerCase().contains("BAN".toLowerCase()))
					{
						ban_No = as.substring(as.indexOf(':')+1,as.lastIndexOf('.')).trim();
						LaunchFlow(flow_Name,ban_No);
					}



					else if (as.toLowerCase().contains("check".toLowerCase()) && as.contains("RR"))
					{
						RR_Name = as.substring(as.indexOf(':')+1,as.lastIndexOf('.')).trim();
						String result=CheckRR(SO_Name,RR_Name);

						if(result.contains("Success"))
						{
							reportingStatus.setCellValue("Pass");

						}
						else
						{
							reportingStatus.setCellValue("Fail");

						}
						reportingReason.setCellValue(result);
						flowEnds();
					}

					 */					

					//////////////////////////////////    CHECK SO   /////////////////////////////////////////////////				

					else if (as.toLowerCase().contains("open".toLowerCase()) && as.toLowerCase().contains("Workflow".toLowerCase())  &&as.toLowerCase().contains("add".toLowerCase()) &&as.toLowerCase().contains("SO".toLowerCase())  )
					{
						SO_Name = as.substring(as.indexOf(':')+1,as.lastIndexOf('.')).trim();
						openInteractive();
						ChangeTab();
						openSO(SO_Name);							 							 
					}
					else if (as.toLowerCase().contains("verify".toLowerCase()) && as.toLowerCase().contains("inputs".toLowerCase()) && as.toLowerCase().contains("visible".toLowerCase()) && as.toLowerCase().contains("selecting".toLowerCase()) )
					{
						HashMap <String,String> excelInputParameter = new HashMap<String, String>();
						String tempVariable = as.substring(as.indexOf(':')+1,as.lastIndexOf('.')).trim();
						String a[] = tempVariable.split(";");
						for(String s :a)
						{
							excelInputParameter.put(s.split("##")[0].toLowerCase(),s.split("##")[1].toLowerCase());
						}
						String result = compareParameters(excelInputParameter, getVariableHashMap());

						if (result.equals("success"))
						{
							reportingStatus.setCellValue("Pass");
						}
						else
						{
							reportingStatus.setCellValue("Fail");

						}
						reportingReason.setCellValue(result);
						//flowEnds();

					}
				}
				if (reportingStatus.getStringCellValue().equals(""))
				{
					reportingStatus.setCellValue("No Run");
				}
			}
			else if (cell !=null && cellIsExecute==null)
			{
				reportingTestName.setCellValue(cellTestName.toString());
				reportingStatus.setCellValue("No Run");
			}
			else if (cell !=null && cellIsExecute==null)
			{
				if(driver!=null)
				{
					driver.close();
					driver.quit();
				}
				
	/*******************   ADDING LAST ROW    ***************************/			
//				row2 = reportingSheet.createRow(i++);	
//				XSSFName name1 = wb.createName();
//				name1.setNameName("runStepLastRow");
//				name1.setRefersToFormula("Sheet2!$D$"+i);
//				reportingStatus = row2.createCell(3);
//				reportingStatus.setCellValue("Last Row");
				FileOutputStream out=new FileOutputStream(excel);
				wb.write(out);
				out.close();
				break;
			}
		}
	}

	private List<WebElement> findElements(By cssSelector) {
		// TODO Auto-generated method stub
		return null;
	}

	public HashMap<String,String> getVariableHashMap()
	{
		HashMap <String,String>variableHm = new HashMap<String, String>();
		List<WebElement> inputVariables = driver.findElements(By.xpath("//*[@class='x-grid3-col x-grid3-cell x-grid3-td-0 x-grid3-cell-first ']"));
		Iterator<WebElement> i = inputVariables.iterator();
		while(i.hasNext())
		{
			WebElement w = i.next();
			if (w.getAttribute("title").contains("must set a non-empty value"))
			{
				variableHm.put(w.getText().toLowerCase(),"mandatory");
			}
			else
			{
				variableHm.put(w.getText().toLowerCase(),"optional");
			}
		}
		return variableHm;
	} 

	public void flowEnds() throws InterruptedException
	{
		System.out.println("End Time : ");
		endTime=getTime();
		reportingEndTime.setCellValue(endTime);
		imageNameCounter=0;

		String subWindowHandler = null;
		Set <String>handles = driver.getWindowHandles();
		//           String currentWindow = driver.getWindowHandle();
		Iterator<String> iterator = handles.iterator();
		while (iterator.hasNext())
		{
			subWindowHandler = iterator.next();
			driver.switchTo().window(subWindowHandler);
			if (driver.getTitle().contains("Workflow Editor"))
			{
				driver.close();
				break;
			}
		}
		ChangeTab();
		//System.out.println(driver.getTitle());
	}


	public WebDriver OpenWFBUrl(String url) {
		ChromeOptions options = new ChromeOptions();
		options.setBinary(outReportingPath+"GoogleChromePortable\\Chrome.exe");
		File file = new File(outReportingPath+"chromedriver.exe");
		System.setProperty("webdriver.chrome.driver", file.getAbsolutePath());
		driver = new ChromeDriver(options);
		driver.get(url);
		driver.getTitle();
		wait = new WebDriverWait(driver, 120000);
		return driver;
	}



	public void Login(String userName, String password) throws IOException {
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/form[1]/center/table/tbody/tr[2]/td/center/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/input")));
		driver.findElement(By.xpath("/html/body/form[1]/center/table/tbody/tr[2]/td/center/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/input"))
		.clear();
		driver.findElement(By.xpath("/html/body/form[1]/center/table/tbody/tr[2]/td/center/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/input"))
		.sendKeys(userName);
		driver.findElement(By.xpath("/html/body/form[1]/center/table/tbody/tr[2]/td/center/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td/input"))
		.clear();
		driver.findElement(By.xpath("/html/body/form[1]/center/table/tbody/tr[2]/td/center/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td/input"))
		.sendKeys(password);
		captureScreen(driver,image_Name);

		driver.findElement(By.xpath("/html/body/form[1]/center/table/tbody/tr[2]/td/center/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[5]/td/input"))
		.click();
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"srv_successok\"]/input")));
		driver.findElement(By.xpath("//*[@id=\"srv_successok\"]/input"))
		.click();
	}

	public void SearchWorkFlow(String workFlowName) throws InterruptedException	
	{
		//		wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//*[@id=\"ext-gen66\"]")));
		//		driver.findElement(By.xpath("//*[@id=\"ext-gen66\"]")).click();
		//		wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//*[@id=\"ext-gen83\"]")));
		//		driver.findElement(By.xpath("//*[@id=\"ext-gen83\"]")).click();
		//		wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//*[@class=\" repository_iconview\"]")));

		wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//button[@class='x-btn-text some_class_that_does_not_exist_but_fixes-rendering repository_ext_btn_align_center'][text()='Change List View']")));
		driver.findElement(By.xpath("//button[@class='x-btn-text some_class_that_does_not_exist_but_fixes-rendering repository_ext_btn_align_center'][text()='Change List View']")).click();
		wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.linkText("Table")));
		driver.findElement(By.linkText("Table")).click();
		Thread.sleep(2000);
		driver.findElement(By.xpath("//*[@id=\"contentTf\"]"))
		.clear();
		driver.findElement(By.xpath("//*[@id=\"contentTf\"]"))
		.sendKeys(workFlowName);			
		wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//button[@class='x-btn-text'][contains(text(),'Filter')]")));		
		driver.findElement(By.xpath("//button[@class='x-btn-text'][contains(text(),'Filter')]")).click();
		Thread.sleep(3000);
		driver.findElement(By.xpath("//button[@class='x-btn-text'][contains(text(),'Filter')]")).click();
	}


	public void OpenWorkFlow(String workFlowName) throws InterruptedException, IOException
	{             
		/*List<WebElement> dw =  driver.findElements(By.xpath("//div[@class='x-panel-body x-panel-body-noheader x-panel-body-noborder']/*[contains(@id,'ext-gen')]"));
                    Iterator<WebElement> i = dw.iterator();
                    while(i.hasNext())
                    {
                                    WebElement w = i.next();
                                    System.out.println("Text : "+w.getText() +" Class : "+w.getClass()+ " Location : "+w.getLocation() +" TagName : "+w.getTagName());


                                                    w.click();
                                                    w.click();
                    }*/
		//		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//img[contains(@title, '" + workFlowName + "')]")));
		//		Thread.sleep(3000);contains(@title, '" + workFlowName + "')
		//		WebElement web = driver.findElement(By.xpath("//*[contains(@title, '" + workFlowName + "')and contains(@class,'title')]"));         
		captureScreen(driver,image_Name);
		wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//div[@class='x-grid3-cell-inner x-grid3-col-title'][text()='"+workFlowName+"']")));
		WebElement workflow=driver.findElement(By.xpath("//div[@class='x-grid3-cell-inner x-grid3-col-title'][text()='"+workFlowName+"']"));
		//		action.doubleClick(workflow).perform();		
		//		WebElement web = driver.findElement(By.xpath("//img[contains(@title, '" + workFlowName + "')]"));
		Actions action = new Actions(driver);
		Action doubleClick = action.doubleClick(workflow).build();
		doubleClick.perform();
		


	}
	
	
	public void ErrorDescripterHandling1() throws InterruptedException{
		
		if(driver.findElements(By.cssSelector("div[class$='x-window-draggable']")).size()!=0){
			int size=driver.findElements(By.cssSelector("div[class$='x-window-draggable']")).size();
			for(int s=0;s<=size-1;s++){
				String ClassName=driver.findElements(By.cssSelector("div[class$='x-window-draggable']")).get(s).findElement(By.cssSelector("span[class='x-window-header-text']")).getText();
				if(ClassName.matches("Error")){
				
			       driver.findElements(By.cssSelector("div[class$='x-window-draggable']")).get(s).findElement(By.cssSelector("div[class$='x-tool-close']")).click();
				}
			}
		}
	}
	
	public void SaveInputs(boolean temp) throws InterruptedException
	{
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(@class,\"x-btn-text\")and contains(@title,'Save')]")));
		driver.findElement(By.xpath("//*[contains(@class,\"x-btn-text\")and contains(@title,'Save')]")).click();
		List<WebElement> dw =driver.findElements(By.xpath("//button[@class=\"x-btn-text\"]"));
		Iterator<WebElement> i = dw.iterator();
		String tempValue ;
		WebElement Tabb1=driver.findElement(By.cssSelector("table[class$='x-btn-wrap x-btn']"));
		 List<WebElement> dw1 =Tabb1.findElements(By.cssSelector("button[class='x-btn-text']"));
			int count3=dw1.size();
			Thread.sleep(1000);
			System.out.println(count3);
		 for(int r=0;r<=count3-1;r++){
			 Thread.sleep(1000);
			 String S112=driver.findElement(By.cssSelector("table[class='x-btn-wrap x-btn']")).findElements(By.cssSelector("button[class='x-btn-text']")).get(r).getText().trim();
			 System.out.println(S112);
			 if (S112.matches("Save")){
				 driver.findElement(By.cssSelector("table[class='x-btn-wrap x-btn']")).findElements(By.cssSelector("button[class='x-btn-text']")).get(r).click();
				 Thread.sleep(1000);
				 break;
				 
			 }
			 
		 }
	}
	public void OpenAutoSave(boolean temp)
	{
		List<WebElement> dw =driver.findElements(By.xpath("//button[@class=\"x-btn-text\"]"));
		Iterator<WebElement> i = dw.iterator();
		String tempValue ;
		if (temp)
			tempValue="Yes";
		else
			tempValue="No";
		while(i.hasNext())
		{
			WebElement w = i.next();
			//System.out.println("Button Name: " +w.getText());
			if (w.getText().equals(tempValue))
			{             
				w.click();
				break;
			}
		}
	}

	public void BreakLock(boolean temp)
	{
		List<WebElement> dw =driver.findElements(By.xpath("//button[@class=\"x-btn-text\"]"));
		Iterator<WebElement> i =dw.iterator();
		String tempValue ;
		if (temp)
			tempValue="OK";
		else
			tempValue="Cancel";
		while(i.hasNext())
		{
			WebElement w = i.next();
			//System.out.println("Button Name: " +w.getText());
			if (w.getText().equals(tempValue))
			{             
				w.click();
				break;
			}
		}
	}
	public void BreakLock1(boolean temp)
	{
		List<WebElement> dw =driver.findElements(By.xpath("//button[@class=\"x-btn-text\"]"));
		Iterator<WebElement> i =dw.iterator();
		String tempValue ;
		if (temp)
			tempValue="Ok";
		else
			tempValue="Cancel";
		while(i.hasNext())
		{
			WebElement w = i.next();
			//System.out.println("Button Name: " +w.getText());
			if (w.getText().equals(tempValue))
			{             
				w.click();
				break;
			}
		}
	}

	public void LaunchFlow(String workFlowName,String  ban_parameters) throws InterruptedException, IOException 
	{             
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(@class,\"x-btn-text\")and contains(@title,'Opens a panel to test workflows with sample inputs.')]")));
		captureScreen(driver,image_Name);
		driver.findElement(By.xpath("//*[contains(@class,\"x-btn-text\")and contains(@title,'Opens a panel to test workflows with sample inputs.')]")).click();
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"workflowTestPanel\"]")));
		SwitchFrame("//*[@id=\"workflowTestPanel\"]");
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"runAsRoles\"]")));
		//captureScreen(image_Name);
		Thread.sleep(3000);
		ErrorDescripterHandling1();
		WebElement web = driver.findElement(By.xpath("//*[@id=\"runAsRoles\"]"));
		web.sendKeys(Keys.chord(Keys.CONTROL, "a"));	
		ErrorDescripterHandling1();

		/*                          
                    List listOpts = web.findElements(By.tagName("option"));
                    int lastIndex = listOpts.size()-1;*/
		//		Select selection = new Select(web);
		//		List<WebElement> options=selection.getOptions();
		//		/*for (int j = 0; j <=lastIndex ; j++) {
		//                                                    selection.selectByIndex(j);
		//                                    }*/
		//		for(WebElement option : options)
		//		{
		//			option.click();
		//		}
		int param_counter=0;
		String [] MultiParameters=ban_parameters.split(";");   
		for(String parameter : MultiParameters)
		{
			
		
		parameter=parameter.trim();
		//Enter BAN	
		if(param_counter==0){
		driver.findElement(By.xpath("//*[@id=\"subscriberId\"]/td[2]/input")).clear();
		ErrorDescripterHandling1();
		driver.findElement(By.xpath("//*[@id=\"subscriberId\"]/td[2]/input")).sendKeys(parameter);
		ErrorDescripterHandling1();
		Thread.sleep(500);
		param_counter++;
		}
		
		if(parameter.contains("@"))
		{
			String[] subParameter=parameter.split("@");		
			String ParameterKey=subParameter[0].trim();		
			String ParameterValue=subParameter[1].trim();	
			
			if(ParameterKey.equalsIgnoreCase("Region"))
			{
				//Input Added
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("createInputName")));
				driver.findElement(By.id("createInputName")).sendKeys("com.motive.modelViewPlatform.webapp.Defines.REGION_IND_SESSION_ATTRIBUTE");
				driver.findElement(By.id("createInputButton")).click();
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("com.motive.modelViewPlatform.webapp.Defines.REGION_IND_SESSION_ATTRIBUTE")));
				driver.findElement(By.name("com.motive.modelViewPlatform.webapp.Defines.REGION_IND_SESSION_ATTRIBUTE")).clear();
				driver.findElement(By.name("com.motive.modelViewPlatform.webapp.Defines.REGION_IND_SESSION_ATTRIBUTE")).sendKeys(ParameterValue);
			}
				
			
		}
		
		}
		captureScreen(driver,image_Name);
		driver.findElement(By.xpath("//*[@id=\"launch-area\"]/a/span")).click();
		ErrorDescripterHandling1();

	}

	/////////////////////////////////////    Create Continue Function     ////////////////////////////////////////////	
	public boolean continueClickTillSuccess() throws InterruptedException
	{
		String workDetails="";

		//		driver.switchTo().defaultContent();
		SwitchFrame("//*[@id='flowFrame']");
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[@id='next_button']")));
		Thread.sleep(2000);
		workDetails=driver.findElement(By.id("workflow-area")).getText();	


		while(true)
		{
			if(workDetails.toLowerCase().contains("workflow has ended"))
			{			
				return false;				
			}
			else if(workDetails.toLowerCase().contains("workflow terminated"))
			{
				return true;		
			}
			else
			{
				wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[@id='next_button']")));
				driver.findElement(By.xpath("//button[@id='next_button']")).click();		
				Thread.sleep(2000);
				workDetails=driver.findElement(By.id("workflow-area")).getText();	
				Thread.sleep(2000);
			}
		}

	}


	public void Click_Button(String button_Name) throws IOException {
		wait.until(ExpectedConditions.elementToBeClickable(By.id(button_Name)));
		driver.findElement(By.id(button_Name)).click();
		wait.until(ExpectedConditions.elementToBeClickable(By.id(button_Name)));
		captureScreen(driver,image_Name);
	}



	public void SwitchFrame(String frameName) {
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(frameName)));
		driver.switchTo().frame(driver.findElement(By.xpath(frameName)));
	}

	public void Run_UtilLSCreateAOTSTicket_Sashank() throws IOException {
		SwitchFrame("//*[@id=\"flowFrame\"]");
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("next_button")));
		Click_Button("next_button");
	}
	public void Run_WANTest() throws IOException{
		SwitchFrame("//*[@id=\"flowFrame\"]");
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("next_button")));
		Click_Button("next_button");
	}


	public void Run_A_UFO_GETNRCharge() throws IOException {
		/*	 List<WebElement> dw =  driver.findElements(By.xpath("//*[contains(@id,'button')]"));
		 Iterator<WebElement> i = dw.iterator();
         while(i.hasNext())
         {
          	WebElement w = i.next();
          	System.out.println("Text : "+w.getText() +" Class : "+w.getClass()+ " Location : "+w.getLocation() +" TagName : "+w.getTagName()+"id " +w.getAttribute("id")+"Name :" +w.getAttribute("name"));
         }  */
		SwitchFrame("//*[@id=\"flowFrame\"]");
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("next_button")));
		driver.findElement(By.xpath("//*[@id=\"promptval\"]"))
		.sendKeys("Trouble Management");
		Click_Button("next_button");
		driver.findElement(By.xpath("//*[@id=\"promptval\"]"))
		.sendKeys("All Prod No Svc");
		Click_Button("next_button");
		driver.findElement(By.xpath("//*[@id=\"promptval\"]"))
		.sendKeys("Battery Back Up Failure");
		Click_Button("next_button");
		driver.findElement(By.xpath("//*[@id=\"promptval\"]"))
		.sendKeys("Dispatched on Demand");
		Click_Button("next_button");
		driver.findElement(By.xpath("//*[@id=\"promptval\"]"))
		.sendKeys("OOS");
		Click_Button("next_button");
		Click_Button("next_button");
	}


	public void OpenDictionary () throws InterruptedException
	{
		driver.switchTo().defaultContent();
		SwitchFrame("//*[@id=\"workflowTestPanel\"]");
		driver.findElement(By.xpath("//*[@id=\"dictionary-button\"]/span")).click();
		ChangeTab();
	}
	public String getChildPosition(String parentXPATH,String childName)
	{
		List<WebElement> dw =driver.findElements(By.xpath(parentXPATH));
		Iterator<WebElement> i = dw.iterator();
		int position=0;
		while(i.hasNext())
		{
			WebElement w = i.next();
			position++;
			if (w.getText().equals(childName))
			{
				return (""+position);
			}
		}		
		return("-1");
	}


	public String getChildXPATH(String initialXPATH,String childName) throws InterruptedException{
		Thread.sleep(1500);
		String xpath = initialXPATH + "/ul/li";
		if (childName!="Get_Value")
		{
			String childPosition =getChildPosition(xpath,childName);
			return (xpath+"["+childPosition+"]");
		}
		else
			return xpath;
	}
	public String searchSODict(String soName) throws InterruptedException
	{
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/div[3]/div/ul")));
		String testModuleXpath=getChildXPATH("/html/body/div[3]/div/ul","testModules");
		driver.findElement(By.xpath(testModuleXpath+"/span")).click();
		String soXpath=getChildXPATH(testModuleXpath,soName);
		driver.findElement(By.xpath(soXpath+"/span")).click();
		return soXpath;
	}
	public String CheckRR(String soName,String RR_Name) throws InterruptedException, IOException
	{             

		driver.findElement(By.xpath("//*[@id=\"sidetreecontrol\"]/a[text()='Collapse All']")).click();
		String soResolutionXpath = getChildXPATH(searchSODict(soName),"resolution");
		driver.findElement(By.xpath(soResolutionXpath+"/span")).click();
		String soAnalysisIdXpath = getChildXPATH(soResolutionXpath,"analysisId");
		driver.findElement(By.xpath(soAnalysisIdXpath+"/span")).click();
		try{
			String RRXpath = getChildXPATH(soAnalysisIdXpath,"Get_Value");
			String RRNameFound = driver.findElement(By.xpath(RRXpath+"/span")).getText();
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(RRXpath+"/span")));
			driver.findElement(By.xpath(RRXpath+"/span")).click();
			captureScreen(driver,image_Name);
			RR_Name = "\""+RR_Name+"\"";
			if(RR_Name.equalsIgnoreCase(RRNameFound))
			{
				return "Success : Expected : "+RR_Name + " :: Found : "+ RRNameFound;
				//return "success";
			}
			else
			{
				return "Error : Expected : "+RR_Name + " :: Found : "+ RRNameFound;
				//return "fail";
			}
		}
		catch(NoSuchElementException e)
		{
			return "Error : Expected : "+RR_Name + " :: No Value Found";
		}
		/*                           System.out.println(driver.findElement(By.xpath("/html/body/div["+dictPosition+"]/div/ul/ul/li[46]")).getText());
                    driver.switchTo().window("CDwindow-E2E0B0E7-C715-48BC-B350-694BA282A5FE");
                    driver.findElement(By.name(soName)).click();
                    List<WebElement> dw = driver.findElements(By.xpath("//*[contains(@tagName,'span')]"));
                    driver.findElement(By.linkText("testModules"));
                    driver.findElement(By.xpath("//span[contains(text(),'analysisId')]")).click();
                    driver.findElement(By.xpath("//span[text()='testModules']")).click();
                    driver.findElement(By.xpath("//span[text()='UFOGetNRCharge']")).click();
                    driver.findElement(By.xpath("//span[text()='resolution']")).click();
                    driver.findElement(By.xpath("//span[text()='analysisId']")).click();
                    System.out.println(driver.findElement(By.xpath("//span[@class='value' and @title='String']")).getText());
                    System.out.println(driver.findElement(By.xpath("/html/body/div[3]/div/ul/ul/li[46]/ul/li/ul/li[3]/ul/li/ul/li/span")).getText());*/               
	}

	public String checkDisplayName(String soName, String displayName) throws InterruptedException, IOException 
	{
		String soResolutionXpath = getChildXPATH(searchSODict(soName),"resolution");
		driver.findElement(By.xpath(soResolutionXpath+"/span")).click();
		try{
			String soDPXpath = getChildXPATH(soResolutionXpath,"displayName");
			driver.findElement(By.xpath(soDPXpath+"/span")).click();		
			String DPXpath = getChildXPATH(soDPXpath,"Get_Value");
			String DPNameFound = driver.findElement(By.xpath(DPXpath+"/span")).getText();
			//		if(DPNameFound.contains("."))
			//		{
			//			DPNameFound=DPNameFound.substring(0,DPNameFound.lastIndexOf('.')).trim();
			//		}		
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(DPXpath+"/span")));
			driver.findElement(By.xpath(DPXpath+"/span")).click();		
			captureScreen(driver,image_Name);
			displayName = "\""+displayName+"\"";
			//		if(displayName.contains(DPNameFound.toLowerCase()))
			if(displayName.equalsIgnoreCase(DPNameFound))
			{
				return "Success : Expected : "+displayName + " :: Found : "+ DPNameFound;
				//return "success";
			}
			else
			{
				return "Error : Expected : "+displayName + " :: Found : "+ DPNameFound;
				//return "fail";
			}
		}
		catch(NoSuchElementException e)
		{
			return "Fail : Expected : "+displayName + " :: No Value Found";
		}
	}

	public String checkResult(String soName,String result) throws InterruptedException, IOException
	{
		String soResultXpath = getChildXPATH(searchSODict(soName),"result");
		driver.findElement(By.xpath(soResultXpath+"/span")).click();
		try{
			String resultXpath = getChildXPATH(soResultXpath,"Get_Value");
			String resultFound = driver.findElement(By.xpath(resultXpath+"/span")).getText();
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(resultXpath+"/span")));
			driver.findElement(By.xpath(resultXpath+"/span")).click();
			captureScreen(driver,image_Name);
			result = "\""+result+"\"";
			if(result.equalsIgnoreCase(resultFound))
			{
				return "Success : Expected : "+result + " :: Found : "+ resultFound;
				//return "success";
			}
			else
			{
				return "Error : Expected : "+result + " :: Found : "+ resultFound;
				//return "fail";
			}
		}
		catch(NoSuchElementException e)
		{

			return "error :"+result+" value not found\n";
		}
	}

	public String checkProperties(String soName, HashMap<String, String> propertyMap ) throws InterruptedException, IOException
	{
		String PropXpath = getChildXPATH(searchSODict(soName),"properties");
		String result="";
		driver.findElement(By.xpath(PropXpath+"/span")).click();
		Set<String> keys=propertyMap.keySet();

		for(String propertyName:keys)
		{
			String propertyXpath= getChildXPATH(PropXpath,propertyName);
			if (propertyXpath==PropXpath+"[-1]")
			{
				result+="Fail : "+propertyName+" not found\n";
			}
			else
			{
				driver.findElement(By.xpath(propertyXpath+"/span")).click();
				try
				{
					String propertyresultXpath = getChildXPATH(propertyXpath,"Get_Value");
					String propertyStatus  = driver.findElement(By.xpath(propertyresultXpath+"/span")).getText();
					wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(propertyresultXpath+"/span")));
					driver.findElement(By.xpath(propertyresultXpath+"/span")).click();
					//					((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(false);", element);
					//					element.click();
					//				((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", propertyStatus);
					//				((Locatable)propertyStatus).getLocationOnScreenOnceScrolledIntoView();
					//				((JavascriptExecutor)driver).executeScript("scrollTo(0,3000)");

					captureScreen(driver,image_Name);
					String propertyValue="\""+propertyMap.get(propertyName)+"\"";
					if(propertyValue.equalsIgnoreCase(propertyStatus))
					{
						result+="PASS :"+propertyName+" expected value : "+propertyStatus+" found\n";
					}
					else
					{
						result+="Fail :"+propertyName+" expected value : "+propertyStatus+", Found Value : "+propertyValue+"\n";
					}
				}
				catch(NoSuchElementException e)
				{
					result+="fail :"+propertyName+" value not found\n";
				}				
			}
		}

		//		String AccstatusXpath= getChildXPATH(PropXpath,propertyName);
		//		if (AccstatusXpath==PropXpath+"[]")
		//		{
		//			return "Error :" + propertyName +" Not Found";
		//		}
		//		driver.findElement(By.xpath(AccstatusXpath+"/span")).click();
		//		String accountstatus  = driver.findElement(By.xpath(AccstatusXpath+"/span")).getText();
		//		if(status.equalsIgnoreCase(accountstatus))
		//		{
		//			return "Success : Expected : "+status + " :: Found : "+ accountstatus;
		//			//return "success";
		//		}
		//		else
		//		{
		//			return "Error : Expected : "+status + " :: Found : "+ accountstatus;
		//			//return "fail";
		//		}
		return result;
	}


	public String checkError(String soName) throws InterruptedException, IOException
	{
		String soErrorXpath = getChildXPATH(searchSODict(soName),"errors");
		wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath(soErrorXpath+"/span")));
		driver.findElement(By.xpath(soErrorXpath+"/span")).click();
		String ErrorXpath = getChildXPATH(soErrorXpath,"Get_Value");		
		try {			
			String errorFound = driver.findElement(By.xpath(ErrorXpath+"/span")).getText();
			if(errorFound==""||errorFound==null)
			{

				captureScreen(driver,image_Name);
				return "Success : No error Found : ";
				//return "success";
			}
			else
			{
				driver.findElement(By.xpath(ErrorXpath+"/span")).click();
				captureScreen(driver,image_Name);
				return "Error : Error Found : "+ errorFound;
				//return "fail";
			}
		} 
		catch (NoSuchElementException e) {
			captureScreen(driver,image_Name);
			return "Success : No error Found : ";
		}
	}



	/*public void captureScreen(String fileName) {
		try {//
			Rectangle screenRect = new Rectangle(Toolkit.getDefaultToolkit().getScreenSize());
			BufferedImage capture = new Robot().createScreenCapture(screenRect);
			//ImageIO.write(capture, "png", new File(fileName));
			imageNameCounter++;
			image_Name = outputPath+"\\"+testName+"\\"+imageNameCounter+".png";
		//} 
		//catch (IOException ex) {
			System.out.println(ex);
		} 
		catch (AWTException ex) {
			System.out.println(ex);
		//}
	//}*/

	public void openSO(String SOName) throws InterruptedException
	{
		wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.linkText("Flow Control")));
		WebElement servOperation= driver.findElement(By.linkText("Flow Control"));
		servOperation.click();
		wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.linkText("Service Operations")));
		WebElement servOperation1= driver.findElement(By.linkText("Service Operations"));
		servOperation1.click();
		wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.className("ORYX_Editor")));
		WebElement droppable= driver.findElement(By.className("ORYX_Editor"));
		wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.linkText(SOName)));
		WebElement dragable=driver.findElement(By.linkText(SOName));
		dragAndDrop(dragable,droppable);
		Thread.sleep(1000);
		//captureScreen(image_Name);
	}


	public void dragAndDrop(WebElement from, WebElement to)
	{
		Actions builder = new Actions(driver);
		Action dragAndDrop = builder.clickAndHold(from)
				.moveToElement(to)
				.release(to)
				.build();

		dragAndDrop.perform();
	}


	public void openInteractive() {
		wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//*[contains(@style,'background-image: url(https://scanr-test3.test.att.com/workflow-builder/images/silk/shape_square_add.png)')and contains(@class,'x-btn-text blist')]")));
		driver.findElement(By.xpath("//*[contains(@style,'background-image: url(https://scanr-test3.test.att.com/workflow-builder/images/silk/shape_square_add.png)')and contains(@class,'x-btn-text blist')]")).click();
		wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//li[@class='x-menu-list-item']")));
		clickListItem("//li[contains(@class,'x-menu-list-item')]", "interactive");	
	}


	public void clickListItem(String ListXPath,String optionName) {
		List<WebElement> options = driver.findElements(By.xpath(ListXPath));
		for (WebElement option : options) {
			//System.out.println(option.getText());
			if(option.getText().equalsIgnoreCase(optionName))
			{
				option.click();
				break;
			}		 
		}
	}

	public String compareParameters(HashMap<String, String> excelParameters,HashMap<String, String> pageParameters)
	{
		Set<String> excelSet=excelParameters.keySet();
		String result="";
		int check=0;

		for(String param: excelSet)
		{
			if(pageParameters.containsKey(param.toLowerCase()))
			{
				if(excelParameters.get(param).equalsIgnoreCase(pageParameters.get(param)))
				{
					//System.out.println("equal");
					check++;
				}
				else {
					result="Expected input parameter- "+param +" : "+excelParameters.get(param)+" Actual Input parameters- "+param +" : "+ pageParameters.get(param);
					return result;
				}
			}
			else
			{
				result = "Expected input parameter- "+param + " Not Found in Variable List";
				return result;
			}

			/*if(excelParameters.get(param).equalsIgnoreCase(pageParameters.get(param)))
    				{
    					System.out.println("equal");
    					check++;
    				}
    				else {
    					result="Expected input parameter- "+param +":"+excelParameters.get(param)+" Actual Input parameters- "+pageParameters+;
    					return result;
    				}*/
		}   		
		if(excelSet.size()==check)
		{
			return "success";	
		}			
		else
		{
			return "error";
		}
	}

	public void ChangeTab() throws InterruptedException
	{             
		Thread.sleep(3000);
		String subWindowHandler = null;
		Set <String>handles = driver.getWindowHandles();
		//           String currentWindow = driver.getWindowHandle();
		Iterator<String> iterator = handles.iterator();
		while (iterator.hasNext())
		{
			subWindowHandler = iterator.next();
			driver.switchTo().window(subWindowHandler);
			//		System.out.println(driver.getTitle());

		}
	}
	
	public String checkMultipleProperties(String soName, String Mainproperty) throws InterruptedException, IOException
	{
		driver.findElement(By.xpath("//*[@id=\"sidetreecontrol\"]/a[text()='Collapse All']")).click();
		String result="";
		String finalResult="";
		String MainPropXpath=getChildXPATH(searchSODict(soName),"properties");					
		String [] MultiProperty=Mainproperty.split("!");      
		int morePropCounter=0;

		for(String propertyMain : MultiProperty)
		{
			if(morePropCounter!=0)
			{
				driver.findElement(By.xpath(MainPropXpath+"/span")).click();	
			}
			driver.findElement(By.xpath(MainPropXpath+"/span")).click();	
			String PropXpath = MainPropXpath;
			propertyMain=propertyMain.trim();

			if(propertyMain.contains("@"))
			{
				String[] subProp=propertyMain.split("@");		
				String propGroup=subProp[0].trim();		
				String propKeyValue=subProp[1].trim();	
				String[] propArray=propKeyValue.split("#");		
				HashMap<String, String> propertyMap= new HashMap<String, String>();
				for(String property:propArray)
				{
					String[] keyVal=property.split(",,");
					propertyMap.put(keyVal[0].trim(), keyVal[1].trim());	 
				}

				/*******************************************************************/	
				String[] propMultiGroup=propGroup.split(",");
				String propertyXpath="";

				if(propMultiGroup.length>1)
				{		
					for(int i=1;i<=propMultiGroup.length;i++)			
					{
						String propMGListElement=propMultiGroup[i-1].trim();
						propertyXpath= getChildXPATH(PropXpath,propMGListElement);

						if (propertyXpath==PropXpath+"[-1]")
						{
							result+="Fail : "+propGroup+" not found\n";
							break;
						}
						else if(i==propMultiGroup.length)
						{
							break;
						}
						else
						{
							PropXpath=propertyXpath;
							driver.findElement(By.xpath(propertyXpath+"/span")).click();
						}				
					}
				}		
				else if(propMultiGroup.length==1)
				{
					propertyXpath= getChildXPATH(PropXpath,propMultiGroup[0]);			
				}
				/***********************************/



				if (propertyXpath==PropXpath+"[-1]")
				{
					result+="Fail : "+propGroup+" not found\n";
				}
				else
				{
					driver.findElement(By.xpath(propertyXpath+"/span")).click();
				}


				Set<String> keys=propertyMap.keySet();
				for(String propertyName:keys)
				{
					String subpropertyXpath= getChildXPATH(propertyXpath,propertyName);
					if (subpropertyXpath==propertyXpath+"[-1]")
					{
						result+="Fail : "+propertyName+" not found\n";
					}
					else
					{
						driver.findElement(By.xpath(subpropertyXpath+"/span")).click();
						try
						{
							String propertyresultXpath = getChildXPATH(subpropertyXpath,"Get_Value");
							String propertyStatus1  = driver.findElement(By.xpath(propertyresultXpath+"/span")).getText().trim();
							String propertyStatus=propertyStatus1.replace("\"", "");
							propertyStatus=propertyStatus1.trim();
							wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(propertyresultXpath+"/span")));
							driver.findElement(By.xpath(propertyresultXpath+"/span")).click();
							captureScreen(driver,image_Name);
							driver.findElement(By.xpath(subpropertyXpath+"/span")).click();
							String propertyValue=propertyMap.get(propertyName);
							propertyValue=propertyValue.replace("\"", "");
							propertyValue=propertyValue.trim();
                            if(propertyValue.equalsIgnoreCase(propertyStatus))
							{
								result+="PASS :"+propertyName+" expected value : "+propertyStatus+" found\n";
							}
							else
							{
								result+="Fail :"+propertyName+" expected value : "+propertyValue+", Found Value : "+propertyStatus+"\n";
							}
						}
						catch(NoSuchElementException e)
						{
							result+="fail :"+propertyName+" value not found\n";
						}				
					}				
				}			
			}
			else
			{
				String[] propArray=propertyMain.split("#");		
				HashMap<String, String> propertyMap= new HashMap<String, String>();
				for(String property:propArray)
				{
					String[] keyVal=property.split(",,");
					propertyMap.put(keyVal[0].trim(), keyVal[1].trim());	 
				}
				Set<String> keys=propertyMap.keySet();
				for(String propertyName:keys)
				{
					String subpropertyXpath= getChildXPATH(MainPropXpath,propertyName);
					if (subpropertyXpath==MainPropXpath+"[-1]")
					{
						result+="Fail : "+propertyName+" not found\n";
					}
					else
					{
						driver.findElement(By.xpath(subpropertyXpath+"/span")).click();
						try
						{
							String propertyresultXpath = getChildXPATH(subpropertyXpath,"Get_Value");
							String propertyStatus1  = driver.findElement(By.xpath(propertyresultXpath+"/span")).getText();
							String propertyStatus=propertyStatus1.replace("\"", "");
							propertyStatus=propertyStatus1.trim();
							wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(propertyresultXpath+"/span")));
							driver.findElement(By.xpath(propertyresultXpath+"/span")).click();
							captureScreen(driver,image_Name);
							driver.findElement(By.xpath(subpropertyXpath+"/span")).click();
							String propertyValue=propertyMap.get(propertyName);
							propertyValue=propertyValue.replace("\"", "");
							propertyValue=propertyValue.trim();
							if(propertyValue.equalsIgnoreCase(propertyStatus))
							{
								result+="PASS :"+propertyName+" expected value : "+propertyStatus+" found\n";
							}
							else
							{
								result+="Fail :"+propertyName+" expected value : "+propertyValue+", Found Value : "+propertyStatus+"\n";
							}
						}
						catch(NoSuchElementException e)
						{
							result+="fail :"+propertyName+" value not found\n";
						}				
					}				
				}
			}


			//finalResult+=result;		
			morePropCounter+=1;
		}
		return result;
	}
	
	
	public String checkMultiplePropertiesInitLS(String soName, String Mainproperty) throws InterruptedException, IOException
	{
		driver.findElement(By.xpath("//*[@id=\"sidetreecontrol\"]/a[text()='Collapse All']")).click();
		String result="";
		String finalResult="";
		String MainPropXpath=getChildXPATH(searchSODict(soName),"SubscriberInformation");					
		String [] MultiProperty=Mainproperty.split("!");      
		int morePropCounter=0;

		for(String propertyMain : MultiProperty)
		{
			if(morePropCounter!=0)
			{
				driver.findElement(By.xpath(MainPropXpath+"/span")).click();	
			}
			driver.findElement(By.xpath(MainPropXpath+"/span")).click();	
			String PropXpath = MainPropXpath;
			propertyMain=propertyMain.trim();

			if(propertyMain.contains("@"))
			{
				String[] subProp=propertyMain.split("@");		
				String propGroup=subProp[0].trim();		
				String propKeyValue=subProp[1].trim();	
				String[] propArray=propKeyValue.split("#");		
				HashMap<String, String> propertyMap= new HashMap<String, String>();
				for(String property:propArray)
				{
					String[] keyVal=property.split(",,");
					propertyMap.put(keyVal[0].trim(), keyVal[1].trim());	 
				}

				/*******************************************************************/	
				String[] propMultiGroup=propGroup.split(",");
				String propertyXpath="";

				if(propMultiGroup.length>1)
				{		
					for(int i=1;i<=propMultiGroup.length;i++)			
					{
						String propMGListElement=propMultiGroup[i-1].trim();
						propertyXpath= getChildXPATH(PropXpath,propMGListElement);

						if (propertyXpath==PropXpath+"[-1]")
						{
							result+="Fail : "+propGroup+" not found\n";
							break;
						}
						else if(i==propMultiGroup.length)
						{
							break;
						}
						else
						{
							PropXpath=propertyXpath;
							driver.findElement(By.xpath(propertyXpath+"/span")).click();
						}				
					}
				}		
				else if(propMultiGroup.length==1)
				{
					propertyXpath= getChildXPATH(PropXpath,propMultiGroup[0]);			
				}
				/***********************************/



				if (propertyXpath==PropXpath+"[-1]")
				{
					result+="Fail : "+propGroup+" not found\n";
				}
				else
				{
					driver.findElement(By.xpath(propertyXpath+"/span")).click();
				}


				Set<String> keys=propertyMap.keySet();
				for(String propertyName:keys)
				{
					String subpropertyXpath= getChildXPATH(propertyXpath,propertyName);
					if (subpropertyXpath==propertyXpath+"[-1]")
					{
						result+="Fail : "+propertyName+" not found\n";
					}
					else
					{
						driver.findElement(By.xpath(subpropertyXpath+"/span")).click();
						try
						{
							String propertyresultXpath = getChildXPATH(subpropertyXpath,"Get_Value");
							String propertyStatus1  = driver.findElement(By.xpath(propertyresultXpath+"/span")).getText();
							String propertyStatus=propertyStatus1.replace("\"", "");
							propertyStatus=propertyStatus1.trim();
							wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(propertyresultXpath+"/span")));
							driver.findElement(By.xpath(propertyresultXpath+"/span")).click();
							captureScreen(driver,image_Name);
							driver.findElement(By.xpath(subpropertyXpath+"/span")).click();
							String propertyValue=propertyMap.get(propertyName);
							propertyValue=propertyValue.replace("\"", "");
							propertyValue=propertyValue.trim();
							propertyValue="\""+propertyMap.get(propertyName)+"\"";
							if(propertyValue.equalsIgnoreCase(propertyStatus))
							{
								result+="PASS :"+propertyName+" expected value : "+propertyStatus+" found\n";
							}
							else
							{
								result+="Fail :"+propertyName+" expected value : "+propertyValue+", Found Value : "+propertyStatus+"\n";
							}
						}
						catch(NoSuchElementException e)
						{
							result+="fail :"+propertyName+" value not found\n";
						}				
					}				
				}			
			}
			else
			{
				String[] propArray=propertyMain.split("#");		
				HashMap<String, String> propertyMap= new HashMap<String, String>();
				for(String property:propArray)
				{
					String[] keyVal=property.split(",,");
					propertyMap.put(keyVal[0].trim(), keyVal[1].trim());	 
				}
				Set<String> keys=propertyMap.keySet();
				for(String propertyName:keys)
				{
					String subpropertyXpath= getChildXPATH(MainPropXpath,propertyName);
					if (subpropertyXpath==MainPropXpath+"[-1]")
					{
						result+="Fail : "+propertyName+" not found\n";
					}
					else
					{
						driver.findElement(By.xpath(subpropertyXpath+"/span")).click();
						try
						{
							String propertyresultXpath = getChildXPATH(subpropertyXpath,"Get_Value");
							String propertyStatus1  = driver.findElement(By.xpath(propertyresultXpath+"/span")).getText().trim();
							String propertyStatus=propertyStatus1.replace("\"", "");
							wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(propertyresultXpath+"/span")));
							driver.findElement(By.xpath(propertyresultXpath+"/span")).click();
							captureScreen(driver,image_Name);
							driver.findElement(By.xpath(subpropertyXpath+"/span")).click();
							String propertyValue1=propertyMap.get(propertyName).trim();
							String propertyValue=propertyValue1.replace("\"", "");
							propertyValue="\""+propertyMap.get(propertyName)+"\"";
							if(propertyValue.equalsIgnoreCase(propertyStatus))
							{
								result+="PASS :"+propertyName+" expected value : "+propertyStatus+" found\n";
							}
							else
							{
								result+="Fail :"+propertyName+" expected value : "+propertyValue+", Found Value : "+propertyStatus+"\n";
							}
						}
						catch(NoSuchElementException e)
						{
							result+="fail :"+propertyName+" value not found\n";
						}				
					}				
				}
			}


			finalResult+=result;		
			morePropCounter+=1;
		}
		return finalResult;
	}


	public String getTime()
	{
		Date date=new Date();
		SimpleDateFormat format=new SimpleDateFormat("MM/dd/yyyy hh:mm");
		System.out.println(date);
		return format.format(date.getTime());
	}
	public  String captureScreen(WebDriver driver, String DestFilePath) throws IOException{
		String TS=fn_GetTimeStamp();
		TakesScreenshot tss=(TakesScreenshot) driver;
   	    File srcfileObj= tss.getScreenshotAs(OutputType.FILE);
   	    File DestFileObj=new File(DestFilePath);
   	    FileUtils.copyFile(srcfileObj, DestFileObj);
   	    return DestFilePath;
	}
	
	//Function to get timestamp
	public static  String fn_GetTimeStamp(){
		DateFormat DF=DateFormat.getDateTimeInstance();
		Date dte=new Date();
		String DateValue=DF.format(dte);
		DateValue=DateValue.replaceAll(":", "_");
		DateValue=DateValue.replaceAll(",", "");
		return DateValue;
	}

	public void createHeaderReport(XSSFSheet reportingSheet)
	{
		XSSFRow row2 = reportingSheet.createRow(0);
		row2.createCell(0).setCellValue("Merged Test Cases");
		row2.createCell(1).setCellValue("Track");
		row2.createCell(2).setCellValue("Test Name");
		row2.createCell(3).setCellValue("Status");
		row2.createCell(4).setCellValue("Reason");
		row2.createCell(5).setCellValue("Start Time");
		row2.createCell(6).setCellValue("End Time");
		row2.createCell(7).setCellValue("Result log Loc");
	}


	public static void main(String[] args) throws IOException, InterruptedException{
		AutomateMain auto = new AutomateMain();
//		auto.Read(args[0]);
		auto.Read("D:\\rv00328221\\TECHM\\E Drive Data\\For selenium\\Copy of NonLS_ADSL.xlsx");

		
	}
}














						