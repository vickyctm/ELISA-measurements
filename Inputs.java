import java.util.Locale;
import java.util.Scanner;
//import java.nio.file.Path;
//import java.nio.file.Paths;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
//import java.util.Iterator;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

//import java.util.Map;
//import java.util.HashMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Inputs {
		
		 //This method is used to load the excel files that would be used as inputs for the 
		//program. The parameter of this method is the file_name that the user has provided.
		public static XSSFWorkbook load_excel(String file_name) throws IOException, FileNotFoundException, InvalidFormatException{
			//Creates a File object with the name of the file provided by the user as parameter.
			File excel_file = new File(file_name);
			
			//Creates a fileinputstream instance for the file. This will be passed in
			//parameter of the XSSFWorkbook object during its creation
			FileInputStream fis = new FileInputStream(excel_file);
			
			//XSSFWork is the root object modeling an Excel XLSX file in Apache
			XSSFWorkbook workbook = new XSSFWorkbook(fis);	
			
//			XSSFWorkbook workbook = new XSSFWorkbook(new File(file_name));
			
			return workbook;
		}
		
		
		//This method goes through the whole ref factors file and returns an array with ALL the rf.
		public static int[] reference_factors (XSSFWorkbook workbook,String reference_factors) {
			//gets the sheet where the reference factors are located
			XSSFSheet rf_sheet = workbook.getSheet(reference_factors);
			int count = count_ints (rf_sheet);
			int[] rf_values = new int [count];
		    int i = 0;
		    int standard_row;
		    int type_column;
			
			for(Row row: rf_sheet) {
				for(Cell cell: row) {
					if(cell.getCellType() != CellType.STRING) {
						type_column = cell.getColumnIndex();
						standard_row = cell.getRowIndex();
					    rf_values[i] = read_reference_factors(rf_sheet, standard_row, type_column);
					    i++;
					} 
				}
			}
			return rf_values;	
		}
		
		
		//This method is called by reference_factors to obtain the value in the cell (rf).
		public static int read_reference_factors(XSSFSheet rf_sheet, int standard_row,int type_column) {
			int rf = 0;
		    
		    XSSFCell rf_value = rf_sheet.getRow(standard_row).getCell(type_column);
			rf = (int) rf_value.getNumericCellValue();
			
			return rf;   			
		}
		
		//gets the number of cells in a document that have ints inside 
		//used for reference_factors
		public static int count_ints (XSSFSheet rf_sheet) {
		    int	count = 0;
		    int test = 0;
		    for(Row row: rf_sheet) {
				for(Cell cell: row) {
					test++;
					if(cell.getCellType() == CellType.NUMERIC) {
						count++;
					}
				}	
			}
			return count;
		}
		
		//method used to find the position where the standard is located 
		public static int find_row(XSSFSheet sheet, String cell_content) {
		    for (Row row: sheet) {
		        for (Cell cell: row) {
		            if (cell.getCellType() == CellType.STRING) {
		                if (cell.getRichStringCellValue().getString().trim().equals(cell_content)) {
		                    return row.getRowNum();  
		                }
		            }
		        }
		    }               
		    return 0;
		}
		
		//methods used to find the position where the type is located 
		public static int find_column(XSSFSheet sheet, String cell_content){
			for (Row row: sheet) {
		        for (Cell cell: row) {
		            if (cell.getCellType() == CellType.STRING) {
		                if (cell.getRichStringCellValue().getString().trim().equals(cell_content)) {
		                    return cell.getColumnIndex();  
		                }
		            }
		        }
		    }               
		    return 0;
		}
		
		
		//returns a list with all the names of HPVs present in the reference factors sheet
		public static String[] hpvlst(XSSFWorkbook workbook, String reference_factors) {
			XSSFSheet rf_sheet = workbook.getSheet(reference_factors);  
		    Row hpvRow = rf_sheet.getRow(0);
		    int numCol = hpvRow.getLastCellNum();
		    String [] hpv = new String [numCol-2];
		    String content;
		    
		    //gets the name of all the HPVs
		    for(int j = 1; j <= (numCol-2); j++) {
		        content = hpvRow.getCell(j).getStringCellValue();
		    	hpv[j-1] = content;
		    }
		    return hpv;
		}
		
		//Using the list of HPVs as reference, checks which ones have rf for a certain standard.
		//returns the name of the standard 
		//COUNTER FOR THE HPV LIST. GOES TO THE FIRST PLACE WITH A REFERENCE FACTOR. GETS THE VALUE
		// OF BOTH REFERENCE FACTOR AND NAME OF THE STANDARD
		 public static int[] id_hpv(XSSFWorkbook workbook, String reference_factors, String[] hpv) {
			 XSSFSheet rf_sheet = workbook.getSheet(reference_factors); 
			 int colindex;
			 int count = count_ints (rf_sheet);
			 int[] rf = new int [count];
			 
			//iterates through all the HPVs  
			 for(int i = 0; i < hpv.length; i++) {
				 int num = 0;
				 colindex = find_column(rf_sheet,hpv[i]);
				 for(int rowindex = 0; rowindex <= rf_sheet.getLastRowNum(); rowindex++) {	
					 Row row = rf_sheet.getRow(rowindex);
					 Cell cell = row.getCell(colindex);
					 	if( (cell.getCellType() != CellType.STRING) ) {
					 		rf[num] = (int) cell.getNumericCellValue();
					 		num++;
					 	}
				 }
			 }
			 
			 return rf;
		 }
		 
	
	    public static void main(String[] args) throws FileNotFoundException, IOException, InvalidFormatException{
		
		//Used to get the users input		 
		Scanner in = new Scanner(System.in);
	        in.useLocale(Locale.US);
	        
	        //path to the Excel file with master sheet, control & standard, reference factors, cut-off values
//	        System.out.print("Path to Excel file (Master,control & standard, reference factors, and cut-off values):");
	        System.out.print("Name of the Excel file and extension (.xlsx): ");
	        String input_file = in.nextLine();
	        File input_path = new File(input_file);
	        
	        //JUST FOR TESTING 
//	        if (input_path.exists()) {
//	        	System.out.println("File created");
//	        }else { System.out.println("no file");}
//	        //for getting the actual path relative to your working directory
//	        System.out.println("reading from "+ input_path.getCanonicalPath());
	       
	        //get the names of the master, control, reference, cut-off values.
//	        System.out.println("Provide the name of the sheet that contains the raw data: ");
//	        String master = in.nextLine();
//	        System.out.println("Provide the name of the sheet that contains the control & standards: ");
//	        String standards = in.nextLine();
	        System.out.println("Provide the name of the sheet that contains the reference factors: ");
	        String reference_factors = in.nextLine();
//	        System.out.println("Provide the name of the sheet that contains the values for 0-negativity: ");
//	        String cut_off = in.nextLine();
	        
	        //Asks for what values you want to use as parameters for your STAT CHECKING
//	        System.out.println("You can set your correlation, slope, and slope ratio values, if you decide to use the default just write NO");
//	        //Asks for the value for the correlation cut_off
//	        System.out.println("Correlation value has to be more than (default 0.9):");
//	        double correlation_value;
//	        if(in.next().equals("NO")) { 
//	        	 correlation_value = 0.9;
//	        } else correlation_value = in.nextDouble();
//	        //Asks for the value for the slope cut_off
//	        System.out.print("Slope value has to be less than (default -0.4):");
//	        double slope_value;
//	        if(in.next().equals("NO")) { 
//	        	 slope_value = -0.4;
//	        } else slope_value = in.nextDouble();
//	        System.out.println();
//	        //Asks for the value for the slope ratio cut_off
//	        System.out.print("Slope Ratio value has to be more than (default 0.5):");
//	        double slope_ratio_value;
//	        if(in.next().equals("NO")) { 
//	        	slope_ratio_value = 0.5;
//	        } else slope_ratio_value = in.nextDouble();
//	        System.out.println();
//	        
//	        
////	        //path to where the output folder will be created 
////	        System.out.print("Output folder: ");
////	        String output_file = in.nextLine();
////	        Path output_path = Paths.get(output_file);
//	        
//	        //number of dilutions. Used to set the size of the list of different dilutions' values
//	        System.out.println("How many different dilutions did you perform: ");
//	        int dnum = in.nextInt();
//	       
//	       //array with all the different dilutions' values
//	        Double[] dilution_list = new Double[dnum];
//	        for (int i = 0; i < dnum; i++) {
//	            System.out.println("Dilution " + (i+1) + ":");
//	                dilution_list[i]= in.nextDouble();
//	        }
//	        System.out.println();
	 
////////////////////////||||||||||||||||||||||||||||//////////////////////	        
	      
	        
            FileInputStream fis = new FileInputStream(input_path);	
			//XSSFWork is the root object modeling an Excel XLSX file in Apache
			XSSFWorkbook input = new XSSFWorkbook(fis);	
	        
//	        //loads the input excelfile
//	        String name = input_path.getName();	  
//	        XSSFWorkbook input = load_excel(name);
	        
			//extracts the rf needed for the calculations
			int [] rf = reference_factors (input,reference_factors);
			for (int i = 0; i < rf.length; i++) {
	            System.out.println(rf[i]);
	        }
	        System.out.println();
	        
	        
	        String [] hpv = hpvlst(input,reference_factors);
	        for(int i = 0; i < hpv.length; i++) {
	        		System.out.println(hpv[i]);
	        }
	        System.out.println();
	        
	       int[] rff= id_hpv(input,reference_factors,hpv);
	       for (int i = 0; i < rff.length; i++) {
	            System.out.println(rff[i]);
	        }
	        System.out.println();
	        
	}
}
