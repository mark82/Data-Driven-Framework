package com.datadriven;



//package com.example.tests;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import com.thoughtworks.selenium.DefaultSelenium;
import com.thoughtworks.selenium.SeleneseTestCase;

@SuppressWarnings("deprecation")
public class DataDriven_Framework extends SeleneseTestCase {
	private static final String String = null;
	// Global variables - accessed any where in this class
	public String vExp, vInc, vLoan;
	public double vTExp, vTInc, vTLoan;
	public String[][] xData; 
	public int xRows, xCols;
	@Before
	public void setUp() throws Exception {
		myprint("Now creating our Selenium object");
		// String xPath = "C:\Selenium\Jul5\csl-data.xls"; Use forward slash instead of backward for file path  
		String xPath = "C:/Selenium/Jul5/csl_data.xls";
		xlRead(xPath); // Reading the excel data
		myprint("XL data read and the rows are " + xRows);

		//selenium = new DefaultSelenium("localhost", 1235, "*iehta", "http://www.chasestudentloans.com/");
		selenium = new DefaultSelenium("localhost", 1236, "*chrome", "http://www.chasestudentloans.com/");
		myprint("Now launching Selenium Log and App browser");
		selenium.start();

	}

	@Test
	public void testCsl1() throws Exception {
		String vTui, vFood, vCC; // Declare Expense variables
		String vJob, vGrants; // Declare income variables
		int vET, vIT, vLT, vR; // Col numbers for the output values
		vET = 11;
		vIT = 12;
		vLT = 13;
		vR = 14; 
		myprint("Starting my main TEST");
		for (int i = 1; i < xRows; i++ ){
			myprint("Row# " + i + " being executed");
			// Read from each row in the XL into variables
			vTui = xData[i][1]; 
			vFood = xData[i][4]; // Taking value into a new variable called vTui
			vCC = xData[i][7];
			vJob = xData[i][8];
			vGrants = xData[i][9];
			myprint("Tui is " + vTui + " Job is " + vJob);
		
			// Simulate the application functionality
			vTExp = Double.parseDouble(vTui) + Double.parseDouble(vFood) + Double.parseDouble(vCC);
			vTInc = Double.parseDouble(vJob) + Double.parseDouble(vGrants);
			vTLoan = vTInc - vTExp;
			myprint("Expense total FROM Script is " + vTExp);
			myprint("Income total FROM Script is " + vTInc);
			myprint("Loan FROM Script is " + vTLoan);
			
			// Test Step 1 
			myprint(ts_001("Budget Calculators", "25000"));
			// Test Step 2
			myprint(ts_002(vTui, vFood, vCC));
			// Test Step 3
			myprint(ts_003(vJob, vGrants));
			// Test Step 4
			myprint(ts_004());
			// Test Step 5 - Compare the results
			if (ts_005()== "Pass") {
				xData[i][vR] = "Pass";
				
			} else {
				xData[i][vR] = "Fail";
		}
			
			// Value from the app back into the array
			xData[i][vET] = vExp;
			xData[i][vIT] = vInc;
			xData[i][vLT] = vLoan; 
		}
		myprint("Ending my main TEST");
	}


	@After
	public void tearDown() throws Exception {
		String res_path= "C:/Selenium/Jul5/csl_data1.xls";
		myprint("Stopping the Selenium RC");
		selenium.stop();
		xlwrite(res_path, xData);
		myprint("@ After - all done");
	}
	
	// Our customized Functions
	public void myprint(String mymessage){
		System.out.println(mymessage);
		System.out.println("~~~~~~~~~~~~~");
	}
	
	public String ts_001(String LinkName, String waitTime){
		//Test Step 1 - Open, go to Bud Calc, wait for page to load
		myprint("Now launching Selenium Log and App browser");
		selenium.open("/");
		selenium.click("link=" + LinkName);
		selenium.waitForPageToLoad(waitTime);
		return "TS001 is a Pass";
	}
	public String ts_002 (String vFTui, String vFFood, String vFCC){
		selenium.type("tuition", vFTui);
		selenium.type("food", vFFood);
		selenium.type("creditcards", vFCC);
		return "TS002 is a Pass";
	}
	public String ts_003(String vFJob, String vFGrants){
		// Income start here
		selenium.type("job", vFJob);
		selenium.type("grants", vFGrants);
		return "TS003 is a Pass";
	}
	public String ts_004(){
		// Clicking on calculate
		selenium.click("//input[@class='calcbutton1']");
		vExp = selenium.getValue("totexp");
		myprint("Expense total is " + vExp);
		vInc = selenium.getValue("totinc");
		myprint("Income total is " + vInc);
		vLoan = selenium.getValue("balance");
		myprint("Loan required is " + vLoan);
		return "TS004 is a Pass";
	}
	public String ts_005(){
		String vE, vI, vL;
		if(vTExp==Double.parseDouble(vExp)){
			myprint("Expenses MATCH");
			vE = "Pass";
		} else {
			myprint("Expenses DO NOT MATCH");
			vE = "Fail";
		}
		if(vTInc==Double.parseDouble(vInc)){
			myprint("Income MATCH");
			vI = "Pass";
		} else {
			myprint("Income DO NOT MATCH");
			vI = "Fail";
		}
		if(vTLoan==Double.parseDouble(vLoan)){
			myprint("Loan MATCH");
			vL = "Pass";
		} else {
			myprint("Loan DO NOT MATCH");
			vL = "Fail";
		}
		if (vE=="Pass" && vI=="Pass" && vL=="Pass"){
			return "Pass";
		} else {
			return "Fail";
		}
	}

	public void xlRead(String sPath) throws Exception{
		File myxl = new File(sPath);
		FileInputStream myStream = new FileInputStream(myxl);
		
		HSSFWorkbook myWB = new HSSFWorkbook(myStream);
		//HSSFSheet mySheet = new HSSFSheet(myWB);
		//HSSFSheet mySheet = myWB.getSheetAt(0);	// Referring to 1st sheet
		HSSFSheet mySheet = myWB.getSheetAt(2);	// Referring to 3rd sheet
		//int xRows = mySheet.getLastRowNum()+1;
		//int xCols = mySheet.getRow(0).getLastCellNum();
		xRows = mySheet.getLastRowNum()+1;
		xCols = mySheet.getRow(0).getLastCellNum();
		myprint("Rows are " + xRows);
		myprint("Cols are " + xCols);
		//String[][] xData = new String[xRows][xCols];
		xData = new String[xRows][xCols];
      for (int i = 0; i < xRows; i++) {
	           HSSFRow row = mySheet.getRow(i);
	            for (int j = 0; j < xCols; j++) {
	               HSSFCell cell = row.getCell(j); // To read value from each col in each row
	               String value = cellToString(cell);
	               xData[i][j] = value;
	               //System.out.print(value);
	               //System.out.print("@@");
	               }
	            //System.out.println("");
	            
	        }	
		
	}
	
	public static String cellToString(HSSFCell cell) {
	// This function will convert an object of type excel cell to a string value
      int type = cell.getCellType();
      Object result;
      switch (type) {
          case HSSFCell.CELL_TYPE_NUMERIC: //0
              result = cell.getNumericCellValue();
              break;
          case HSSFCell.CELL_TYPE_STRING: //1
              result = cell.getStringCellValue();
              break;
          case HSSFCell.CELL_TYPE_FORMULA: //2
              throw new RuntimeException("We can't evaluate formulas in Java");
          case HSSFCell.CELL_TYPE_BLANK: //3
              result = "-";
              break;
          case HSSFCell.CELL_TYPE_BOOLEAN: //4
              result = cell.getBooleanCellValue();
              break;
          case HSSFCell.CELL_TYPE_ERROR: //5
              throw new RuntimeException ("This cell has an error");
          default:
              throw new RuntimeException("We don't support this cell type: " + type);
      }
      return result.toString();
  }
	public void xlwrite(String xlPath, String[][] xldata) throws Exception {
		System.out.println("Inside XL Write");
  	File outFile = new File(xlPath);
      HSSFWorkbook wb = new HSSFWorkbook();
         // Make a worksheet in the XL document created
      /*HSSFSheet osheet = wb.setSheetName(1,"TEST");*/
      HSSFSheet osheet = wb.createSheet("TESTRESULTS");
      // Create row at index zero ( Top Row)
  	for (int myrow = 0; myrow < xRows; myrow++) {
  		//System.out.println("Inside XL Write");
	        HSSFRow row = osheet.createRow(myrow);
	        // Create a cell at index zero ( Top Left)
	        for (int mycol = 0; mycol < xCols; mycol++) {
	        	HSSFCell cell = row.createCell(mycol);
	        	// Lets make the cell a string type
	        	cell.setCellType(HSSFCell.CELL_TYPE_STRING);
	        	// Type some content
	        	cell.setCellValue(xldata[myrow][mycol]);
	        	//System.out.print("..." + xldata[myrow][mycol]);
	        }
	        //System.out.println("..................");
	        // The Output file is where the xls will be created
	        FileOutputStream fOut = new FileOutputStream(outFile);
	        // Write the XL sheet
	        wb.write(fOut);
	        fOut.flush();
//		    // Done Deal..
	        fOut.close();
  	}
  }
}
