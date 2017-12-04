package Scanr;


import com.jcraft.jsch.*;

import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.BufferedOutputStream;
import java.io.BufferedWriter;
import java.io.Closeable;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.FilterOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.PrintStream;
import java.text.DateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Properties;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Keys;

public class admin_Logs {


	static public ArrayList<String> commands = new ArrayList<String>();

	static public String file;
	static public XSSFWorkbook wb;
	static public XSSFSheet sheet;
	static public XSSFRow row;
	static public XSSFCell cell1, cell2, cell3, cell4, cell5, cell6, cell7, cell8, cell9, cell11,cell12,cell_Result,cell_Result_Apis;
	static public XSSFRow row_dynamic;

	static public String server;
	static public String username;
	static public String logFile;
	static public String logFile2;
	static public String privatekey;
	static public String command1;
	static public String command2;
	static public String command3;
	static public String command4;
	static public String command5;
	static public String Ts;

	public OutputStream runCommands(String ExcelPath, String testName) {

		OutputStream log = null;

		try {	
			file = ExcelPath;
			wb = new XSSFWorkbook(new FileInputStream(file));
			sheet = wb.getSheet("Admin_log");
			row = sheet.getRow(1);
			cell1 = row.getCell(0);
			cell2 = row.getCell(1);
			cell3 = row.getCell(2);
			cell5 = row.getCell(4);
			cell6 = row.getCell(5);
			cell7 = row.getCell(6);
			cell8 = row.getCell(7);
			cell11=row.getCell(10);
			cell12=row.getCell(11);
			

			server = cell1.getStringCellValue();
			username = cell2.getStringCellValue();
			logFile = cell3.getStringCellValue();

			privatekey = cell5.getStringCellValue();
			command1 = cell6.getStringCellValue();
			command2 = cell7.getStringCellValue();
			command3 = cell8.getStringCellValue();
			//command4 = cell11.getStringCellValue();
			//command5 = cell12.getStringCellValue();
			commands.add(command1);
			commands.add(command2);
			commands.add(command3);
			
			DateFormat DF=DateFormat.getDateTimeInstance();
			Date dte=new Date();
			Ts=DF.format(dte);
			Ts=Ts.replaceAll(":", "_");
			Ts=Ts.replaceAll(",","_");
			Ts=Ts.replaceAll(" ","_");
			
			//*****************Creating Session****************************************** 
			JSch js = new JSch();
			Session s = js.getSession(username, "139.76.209.171", 22);
			s.setPassword("qwerty123");
			Properties config = new Properties();
			config.put("StrictHostKeyChecking", "no");
			s.setConfig(config);
			s.connect();
			s.setTimeout(6000);

			Channel channel = s.openChannel("shell");// only shell
			File nfile = new File(testName+".txt");
			log = new BufferedOutputStream(new FileOutputStream(nfile));
			channel.setOutputStream(log);
			PrintStream shellStream = new PrintStream(channel.getOutputStream()); 
			channel.setInputStream(null);
			channel.connect(3000);
			
			//****************Loop to enter commands**************************************
			int i = 0;
			for (String command : commands) {
				if(i==2)
					shellStream.println(command+"| tee /tmp/tmp"+Ts+".txt");	 
				shellStream.println(command);
				shellStream.flush();
				System.out.println(command);
				i=i+1;
			}//***************************End of Loop**************************************

		} catch (Exception e) {
			System.err.println("ERROR: Connecting via shell to ");
			e.printStackTrace();
		}
		return (OutputStream) log;
	}

	public void Disconnect(OutputStream log){
		try {
			log.close();
			
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	

	

	//***************************** For Awk Commands Execution*******************************
	public String runCommands_s2(String testName,int i,String ban_No_Parameters,String Interface_Api_Name_Status) {
		StringBuilder line = new StringBuilder();
		String Log = null;
		try{

			//file = "D:\\E Drive Data\\For selenium\\AdminLog_Automation.xlsx";
			//row = sheet.getRow(1);
			row_dynamic = sheet.createRow(i+3);
			row_dynamic = sheet.getRow(i+3);
			
			cell1 = row_dynamic.createCell(0);
			cell1.setCellValue(testName);
			
			cell2 = row_dynamic.createCell(1);
			cell2 = row_dynamic.getCell(1);
			cell3=row_dynamic.createCell(2);
			cell3=row_dynamic.getCell(2);
			
			cell4 = row.getCell(3);
			logFile2 = cell4.getStringCellValue();
			
			cell4=row_dynamic.createCell(3);
			cell4=row_dynamic.getCell(3);
			
		//*************************Creating Session for Awk Command****************************
			JSch js2 = new JSch();
			Session s2 = js2.getSession(username, "139.76.209.171", 22);
			s2.setPassword("qwerty123");
			Properties config2 = new Properties();
			config2.put("StrictHostKeyChecking", "no");
			s2.setConfig(config2);

			s2.connect();
			Channel channel2 = s2.openChannel("shell");// only shell
			//OutputStream log2 = new BufferedOutputStream(new FileOutputStream(logFile2));
			//channel2.setOutputStream(log2);

			PrintStream shellStream2 = new PrintStream(channel2.getOutputStream()); // printStream
			channel2.setInputStream(null);

			Thread.sleep(1000);
			
			//************************Reading Data from server*********************************************
			InputStream inStream = channel2.getInputStream();
			channel2.connect(3000);

			shellStream2.println(command1);
			shellStream2.flush();
			System.out.println(command1);

			shellStream2.println(command2);
			shellStream2.flush();
			System.out.println(command2);

			//Command to fetch occurrence of #### in the logs
			shellStream2.println("grep -n -B11 "+ban_No_Parameters+" /tmp/tmp"+Ts+".txt | head -1 | cut -d \"-\" -f1");
			shellStream2.flush();
			System.out.println("grep -n -B11 "+ban_No_Parameters+" /tmp/tmp"+Ts+".txt | head -1 | cut -d \"-\" -f1");
			
			//Reading Starting Line Number extracted after running the above command
			InputStreamReader tout=new InputStreamReader(inStream);
			char toAppend = ' ';
			Thread.sleep(1000);
			while (tout.ready()) {
				toAppend = (char) tout.read();
				line.append(toAppend);}
			System.out.print(line.toString());
			
			String startLineNum;String[] str;
			int startLineNum_size=line.toString().split("\n").length;
			str=line.toString().split("\n");
			startLineNum=str[startLineNum_size-3].trim();
					
			//Command to fetch End Line Number 
			shellStream2.println("cat -n /tmp/tmp"+Ts+".txt | sed -n '/"+ban_No_Parameters+"/,/}>/p' | tail -1 | awk '{print $1};'");
			shellStream2.flush();
			System.out.println("cat -n /tmp/tmp"+Ts+".txt | sed -n '/"+ban_No_Parameters+"/,/}>/p' | tail -1 | awk '{print $1};'");
			
			//Reading End Line Number extracted after running the above command
			tout = null;
			tout=new InputStreamReader(inStream);
			toAppend = ' ';
			Thread.sleep(1000);
			while (tout.ready()) {
				toAppend = (char) tout.read();
				line.append(toAppend);}
			System.out.print(line.toString());
			String endLineNum;
			int endLineNum_size=line.toString().split("\n").length;
			str=line.toString().split("\n");
			endLineNum=str[endLineNum_size-3].trim();
					
			
			//Commands to get logs between extracted line numbers
			shellStream2.println("sed -n '"+startLineNum+","+endLineNum+"p' /tmp/tmp"+Ts+".txt");
			shellStream2.flush();
			System.out.println("sed -n '"+startLineNum+","+endLineNum+"p' /tmp/tmp"+Ts+".txt");
			
			//Reading Log extracted after running the above command
			tout = null;
			tout=new InputStreamReader(inStream);
			toAppend = ' ';
			Thread.sleep(1000);
			while (tout.ready()) {
				toAppend = (char) tout.read();
				line.append(toAppend);}
			System.out.print(line.toString());
			
			Log=line.toString().trim();
			
			channel2.disconnect();
			s2.disconnect();
			
			//_________________________________________________________________________________________________________________________________________
			s2 = js2.getSession(username, "139.76.209.171", 22);
			s2.setPassword("qwerty123");
			config2.put("StrictHostKeyChecking", "no");
			s2.setConfig(config2);
			
			
			s2.connect();
			channel2 = s2.openChannel("shell");
			inStream = channel2.getInputStream();
			shellStream2 = new PrintStream(channel2.getOutputStream()); // printStream
			channel2.setInputStream(null);
			channel2.connect(3000);
			
			shellStream2.println(command1);
			shellStream2.flush();
			System.out.println(command1);

			shellStream2.println(command2);
			shellStream2.flush();
			System.out.println(command2);
			
			
			
			shellStream2.println("awk '/severity:/{if($33==\"FAILURE\")print\"Failed|\";if($33==\"INFORMATION\")print\"Passed|\";}/call/{print\"|\"$3;}/parameters/{if($2!=\"<?xml\")print\"|\"$0;}' /tmp/tmp"+Ts+".txt > /tmp/newtmp"+Ts+".txt");
			shellStream2.flush();
			System.out.println("awk '/severity:/{if($33==\"FAILURE\")print\"Failed|\";if($33==\"INFORMATION\")print\"Passed|\";}/call/{print\"|\"$3;}/parameters/{if($2!=\"<?xml\")print\"|\"$0;}' /tmp/tmp"+Ts+".txt > /tmp/newtmp"+Ts+".txt");

			shellStream2.println("grep -B2 -A5000 "+ban_No_Parameters+" /tmp/newtmp"+Ts+".txt | tac | grep -A5000"+" "+ban_No_Parameters);
			shellStream2.flush();
			System.out.println("grep -B2 -A5000 "+ban_No_Parameters+" /tmp/newtmp"+Ts+".txt | tac | grep -A5000 "+ban_No_Parameters);

			//*******************Extracting text from server file*********************************************//	
			Thread.sleep(1000);
			tout = null;
			tout=new InputStreamReader(inStream);
			
			toAppend =' ';
			Thread.sleep(1000);
			while (tout.ready()) {
				toAppend = (char) tout.read();
				line.append(toAppend);


			}
			System.out.print(line.toString());
			Thread.sleep(1000);
			

			//*********************Making Final Result from extracted logs*************************************//           
			String line_str = line.toString();
			String[] String_Apis;  

			String_Apis = line_str.split("\n");
			String Result_Apis=" ";
			String Final_Result;
			int size = String_Apis.length;
			String[] Api_Name_Status;
			String Result_Status = "";
			Api_Name_Status = Interface_Api_Name_Status.split("!");
			for (int iterator=0;iterator<size-1;iterator++) {

				if(String_Apis[iterator].contains("parameters:") )
				{

					while (!String_Apis[iterator+1].contains("parameters:") && iterator<String_Apis.length-2)
					{
						Result_Apis = Result_Apis.concat(String_Apis[iterator]);
						Result_Apis = Result_Apis.concat("\n");
						iterator=iterator+1; 

					}
					Result_Apis = Result_Apis.concat(String_Apis[iterator]);
					Result_Apis = Result_Apis.concat("\n");

				}
			
				
			}
			
			//*********************To Check for Pass/Fail of a particular api
			for (int iterator=0;iterator<size-1;iterator++) {
				
				if(String_Apis[iterator].contains(Api_Name_Status[0]))
				{
					if(String_Apis[iterator+1].contains("Pass"))
						Result_Status = Api_Name_Status[0]+"_"+"Pass";
					else
						Result_Status = Api_Name_Status[0]+"_"+"Fail";
				}	
			}
			
			
			if(Result_Apis.contains("Fail"))
				Final_Result="Fail";
			else
				Final_Result="Pass";
			//****************************************************************************************************//

			//			byte[] buffer = new byte[11024];																		
			//			InputStream in=channel2.getInputStream();
			//			String line = "";
			//		      
			//		            while (in.available() > 0) {
			//		                int j = in.read(buffer, 0, 11024);
			//		             
			//		                if (j < 0) {
			//		                    break;
			//               }
			//		                line = new String(buffer, 0, j);
			//		                System.out.println(line);
			//		            }

			
			//******************************Writing Pass/Fail Api list into logfile2*******************************
			BufferedWriter writer = null;
			writer = new BufferedWriter( new FileWriter(logFile2));
			writer.write(line.toString());
			writer.close( );

			channel2.disconnect();
			s2.disconnect();

			//*****************************Writing Final Result into Excel****************************************
			cell1.setCellValue(testName);
			cell3.setCellValue(Result_Apis);
			cell2.setCellValue(Final_Result);
			cell4.setCellValue(Result_Status.toString());
			
			
		} catch (Exception e) {
			System.err.println("ERROR: Connecting via shell to ");
			e.printStackTrace();
		}
		return Log;

		
	}//***********************************End of Function**********************************************************

	
	
	
	//**********************************************************************************************************************
	//@@@@@@@@@@@@@@@@@@@@@@@@@@@@ Verify Request of any Api you send in Parameters@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
	//**********************************************************************************************************************
	public void Verify_Request_Tags(String Log_Text,int i,String ParameterName_Value,String Interface_Api_Name) {
		try{

			String line_str = Log_Text;
			String[] String_Apis;  
			String_Apis = line_str.split("\n");
			String Request="";
			int size = String_Apis.length;
			
			
		//*************************Extracting Request for the Parameterized Api**********************************************	
			for (int iterator=0;iterator<size-1 && Request=="";iterator++) {
				System.out.println(String_Apis[iterator]);
				if(String_Apis[iterator].contains(Interface_Api_Name) )
				{
					int flag = iterator+2;
					while(!String_Apis[flag].contains("parameters:") && flag<String_Apis.length-2){
						flag=flag+1;
						System.out.println(String_Apis[flag]);
						
					}
						 
					
					while (!String_Apis[flag].contains("time:") && flag<String_Apis.length-2)
					{
						Request = Request.concat(String_Apis[flag]);
						Request = Request.concat("\n");
						flag=flag+1; 
					}
				}
			}//********************End of Request Extraction****************************************************************
			
		String Pv_Array[]=ParameterName_Value.split(";");
		int Pv_size = Pv_Array.length;
		String[] PV_Split;
		String Argument;
		String Request_Status="";
		
		for(int j=0;j<=Pv_size-1;j++)
		{
			PV_Split=Pv_Array[j].split("#");
			Argument = PV_Split[1]+"</";
			Argument=Argument.toString();
			
			if(Request.contains(PV_Split[1]))
			{
				int indexofValue = Request.indexOf(PV_Split[1]);
				String Request1 = Request.substring(indexofValue);
				if(Request.contains(PV_Split[0]))
					{
						int indexofParameter = (Request1.indexOf(">"))+1;
						Request = Request1.substring(0, indexofParameter);
						if(Request.contains(Argument))
							Request_Status=Request_Status.concat(PV_Split[0]+"Value_Exists;");
						else
							Request_Status=Request_Status.concat(PV_Split[0]+"Parameter-Value Does Not Exist;");
					}else
						Request_Status=Request_Status.concat(PV_Split[0]+"Parameter Does Not Exist;");
			}else
				Request_Status=Request_Status.concat(PV_Split[1]+"Parameter's Value Does Not Exist;");
				
		}
		row_dynamic = sheet.getRow(i+3);
		cell5=row_dynamic.createCell(4);
		cell5=row_dynamic.getCell(4);
		cell5.setCellValue(Request_Status.trim());
			
		channel2.disconnect();
		s2.disconnect();
		
		FileOutputStream fout = new FileOutputStream(file);	
		wb.write(fout);
		fout.close();
			
		} catch (Exception e) {
			System.err.println("ERROR: Connecting via shell to ");
			e.printStackTrace();
		}


	}//***********************************End of Function**********************************************************
	
	
	

	public static void main(String[] args) throws IOException,
	InterruptedException {
		admin_Logs auto = new admin_Logs();
		file = "D:\\E Drive Data\\For selenium\\AdminLog_Automation.xlsx";
		OutputStream log=auto.runCommands(file,"tname");
		//******************************Execute Commands you want to execute and then use parameterized code for ban****************
		auto.Disconnect(log);
		auto.runCommands_s2("TestCaseName",1,"444444444","CSIInquireDSLServiceDetail");
	}

}
