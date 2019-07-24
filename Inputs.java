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
	static //Global variables
	String id = null;
    public static int call_counter = 0; //USED TO COUNT HOW MANY TIMES THE id_hpv METHOD WAS CALLED 
    public static int [] dilution;
	
	
	
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
		public static int[] reference_factors (XSSFSheet rf_sheet) {
			int count = count_ints (rf_sheet);
			int[] rf_values = new int [count];
		    int i = 0;
		    int standard_row;
		    int type_column;
			
			for(Row row: rf_sheet) {
				for(Cell cell: row) {
					if(cell.getCellType() == CellType.NUMERIC) {
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
		public static String[] hpvlst(XSSFSheet rf_sheet, String reference_factors) {
		    Row hpvRow = rf_sheet.getRow(0);
		    int numCol = hpvRow.getLastCellNum();
		    String [] hpv = new String [numCol-1];
		    String content;
		    
		    //gets the name of all the HPVs
		    for(int j = 1; j <= (numCol-1); j++) {
		        content = hpvRow.getCell(j).getStringCellValue();
		    	hpv[j-1] = content;
		    }
		    return hpv;
		}
		
		//The HPV list is used as a sort of counter. 
		 public static int id_hpv(XSSFSheet rf_sheet, String[] hpv) {
//			 int master_counter = count_ints (rf_sheet);
			 call_counter++;
			 int rf = 0;
			
			 //THIS IS HOW YOU ITERATE THROUGH A COLUMN!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
				int colindex = find_column(rf_sheet,hpv[(call_counter-1)]);
				 for(int rowindex = 0; rowindex <= rf_sheet.getLastRowNum(); rowindex++) {	
					 Row row = rf_sheet.getRow(rowindex);
					 Cell cell = row.getCell(colindex);
					 	if(cell.getCellType() == CellType.NUMERIC) {
					 		rf = (int) cell.getNumericCellValue();
					 		id = row.getCell(0).getStringCellValue();
					 	}
				 }
			 
//			 int doesitwork = (call_counter-1);
//			 int colindex = find_column(rf_sheet,hpv[(call_counter-1)]);	 
//			 int check;
//			 
//			 	for(int rowindex = 0; rowindex <= rf_sheet.getLastRowNum(); rowindex++) {	
//			 		Row row = rf_sheet.getRow(rowindex);
//			 		Cell cell = row.getCell(colindex);
//				 		if(cell.getCellType() == CellType.NUMERIC) {
//				 			rf = (int) cell.getNumericCellValue();
//				 			id = row.getCell(0).getStringCellValue();
//				 		
//				 			check = (rowindex + 1);
//				 			Row checkrow = rf_sheet.getRow(check);
//				 			Cell checkcell = checkrow.getCell(colindex);
//				 				if(checkcell.getCellType() == CellType.NUMERIC) {
//				 					call_counter--;	
//				 				}
//				 				break;
//				 		}
//			 	}
			 
			 return rf;
		 }
		 
		 //This method checks the number of dilutions there is by comparing the id and run values 
		 public static int size_dilutionlst (XSSFSheet raw_sheet) {
			 //column indexes
			 int RUN = 0;
			 int ID = 1;
			 int DILUTION = 2;
			 
			 int size = 1;
			
			 //Initial cells used to then compared with the others 
			 Row ROW1 = raw_sheet.getRow(2);
			 Cell R1 = ROW1.getCell(RUN);
			 Cell ID1 = ROW1.getCell(ID);
			 
			 for(int rowindex = 2; rowindex <= 10; rowindex++) {	
				 Row row = raw_sheet.getRow(rowindex);
				 Cell r_cell = row.getCell(RUN);
				 Cell id_cell = row.getCell(ID);
				 
				 Cell d_cell = row.getCell(DILUTION);
				 
				 if(((r_cell.getNumericCellValue()) == (R1.getNumericCellValue())) && (id_cell.getStringCellValue().equals(ID1.getStringCellValue()))) {
					 size++;
				 } 
			 }
			 return size;
		 }
		 
		 //Takes care of the array with all the dilutions needed for this raw data sheet
		 public static void get_dilutions (XSSFSheet raw_sheet, int size) {
			 int DILUTION = 2;
			 dilution = new int [size];
	
			 for(int rowindex = 1; rowindex <= size; rowindex++) {	
				 Row row = raw_sheet.getRow(rowindex);
				 Cell d_cell = row.getCell(DILUTION);
				 dilution[(rowindex-1)] = (int) d_cell.getNumericCellValue();
			 }
		 }
		 
		 
		 //Extracts a line of the raw data document
		 public static double [] line_raw (XSSFSheet raw_sheet, String type, int pos, int size) {
			 //column indexes 
			 int type_col = find_column(raw_sheet,type);
			 double [] data = new double [size];

				 for(int rowindex = (pos+1); rowindex < (size+(pos+1)); rowindex++) {	
					 Row row = raw_sheet.getRow(rowindex);
					 Cell data_cell = row.getCell(type_col);
					 data[((rowindex-(pos+1)))] =  data_cell.getNumericCellValue();
				 }
			 return data;
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
	        System.out.println("Provide the name of the sheet that contains the raw data: ");
	        String master = in.nextLine();
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
	      
	       //loads the input excelfile 
            FileInputStream fis = new FileInputStream(input_path);	
			@SuppressWarnings("resource")
			XSSFWorkbook input = new XSSFWorkbook(fis);	  
//	        XSSFWorkbook input = load_excel(input_path.getName());
	        
			//Assigns names to the sheets
			XSSFSheet raw_sheet = input.getSheet(master);
//			XSSFSheet ctrl_sheet = input.getSheet(standards);
			XSSFSheet rf_sheet = input.getSheet(reference_factors);
//			XSSFSheet cutoff_sheet = input.getSheet(cut_off);
			
			
			
			//This is all the processing required for the reference factors sheet.
			//The hpv list -- master counter, the id_hpv to obtain the standard and the rf value.
			
//			int [] rf = reference_factors (input,rf_sheet);   //extracts the rf needed for the calculations
	        String [] hpv = hpvlst(rf_sheet,reference_factors);  //creates the list of all hpvs 
	        int rf= id_hpv(rf_sheet,hpv); //rf factor and id for that specific virus
//            System.out.println(rf);
//            System.out.println(id);

	        
	        //This is all the processing required for a raw data sheet
		    int size = size_dilutionlst (raw_sheet);
		    get_dilutions(raw_sheet,size); //gets the dilution list 
		    double [] data = new double [size];
		    int pos = 0;
		    int ending = (raw_sheet.getLastRowNum() - size);
		   
		    //counter for each line 
		    while (pos <= ending) {
		    	data = line_raw (raw_sheet,"HPV 6",pos,size); 
		   
		    	//all the calculations go here 
		    	
		    	pos = (pos + size);
		    }
		   
		    for(int i = 0; i < data.length; i++) { //for testing
		    System.out.println(data[i]);
		    }
		    
		   
	        
	                 
	}
}
