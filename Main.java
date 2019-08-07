import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Locale;
import java.util.Scanner;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author Victoria Torres
 *
 */


public class Main {
	public static void main(String[] args) throws FileNotFoundException, IOException, InvalidFormatException {
		// Used to get the users input
		Scanner in = new Scanner(System.in);
		in.useLocale(Locale.US);

		// path to the Excel file with master sheet, control & standard, reference
		// factors, cut-off values
		// System.out.print("Path to Excel file (Master,control & standard, reference
		// factors, and cut-off values):");
		System.out.print("Name of the Excel file and extension (.xlsx): ");
		String input_file = in.nextLine();
		File input_path = new File(input_file);

		// JUST FOR TESTING
		// if (input_path.exists()) {
		// System.out.println("File created");
		// }else { System.out.println("no file");}
		// //for getting the actual path relative to your working directory
		// System.out.println("reading from "+ input_path.getCanonicalPath());

		// get the names of the master, control, reference, cut-off values.
		System.out.println("Provide the name of the sheet that contains the raw data: ");
		String master = in.nextLine();
		System.out.println("Provide the name of the sheet that contains the control & standards: ");
		String standards = in.nextLine();
		System.out.println("Provide the name of the sheet that contains the reference factors: ");
		String reference_factors = in.nextLine();
		// System.out.println("Provide the name of the sheet that contains the values
		// for 0-negativity: ");
		// String cut_off = in.nextLine();

		// Asks for what values you want to use as parameters for your STAT CHECKING
		// System.out.println("You can set your correlation, slope, and slope ratio
		// values, if you decide to use the default just write NO");
		// //Asks for the value for the correlation cut_off
		// System.out.println("Correlation value has to be more than (default 0.9):");
		// double correlation_value;
		// if(in.next().equals("NO")) {
		// correlation_value = 0.9;
		// } else correlation_value = in.nextDouble();
		// //Asks for the value for the slope cut_off
		// System.out.print("Slope value has to be less than (default -0.4):");
		// double slope_value;
		// if(in.next().equals("NO")) {
		// slope_value = -0.4;
		// } else slope_value = in.nextDouble();
		// System.out.println();
		// //Asks for the value for the slope ratio cut_off
		// System.out.print("Slope Ratio value has to be more than (default 0.5):");
		// double slope_ratio_value;
		// if(in.next().equals("NO")) {
		// slope_ratio_value = 0.5;
		// } else slope_ratio_value = in.nextDouble();
		// System.out.println();
		//
		//
		//// //path to where the output folder will be created
		//// System.out.print("Output folder: ");
		//// String output_file = in.nextLine();
		//// Path output_path = Paths.get(output_file);
		//
		// //number of dilutions. Used to set the size of the list of different
		// dilutions' values
		// System.out.println("How many different dilutions did you perform: ");
		// int dnum = in.nextInt();
		//
		// //array with all the different dilutions' values
		// Double[] dilution_list = new Double[dnum];
		// for (int i = 0; i < dnum; i++) {
		// System.out.println("Dilution " + (i+1) + ":");
		// dilution_list[i]= in.nextDouble();
		// }
		// System.out.println();

		//////////////////////// ||||||||||||||||||||||||||||//////////////////////

		// loads the input excelfile
		FileInputStream fis = new FileInputStream(input_path);
		@SuppressWarnings("resource")
		XSSFWorkbook input = new XSSFWorkbook(fis);
		// XSSFWorkbook input = load_excel(input_path.getName());

		// Assigns names to the sheets
		XSSFSheet raw_sheet = input.getSheet(master);
		XSSFSheet ctrl_sheet = input.getSheet(standards);
		XSSFSheet rf_sheet = input.getSheet(reference_factors);
		// XSSFSheet cutoff_sheet = input.getSheet(cut_off);

	//////////////////////////////////////////////////////////////////////////////////////////	
    //////////////////////////////////////////////////////////////////////////////////////////	
    //////////////////////////////////////////////////////////////////////////////////////////		
		//creates the output workbook
		//we are going to create this new sheet once we getthe hpv
//			    XSSFWorkbook output = new XSSFWorkbook();
		
		
		// This is all the processing required for the reference factors sheet.
		// The hpv list -- master counter, the id_hpv to obtain the standard and the rf
		// value.

		// int [] rf = reference_factors (input,rf_sheet); //extracts the rf needed for
		// the calculations
		String[] hpv = Inputs.hpvlst(rf_sheet); // creates the list of all hpvs
		double rf = Inputs.id_hpv(rf_sheet, hpv); // rf factor and id for that specific virus
		// System.out.println(rf);
		// System.out.println(id);
		
	
	    
	    
	    //CHECKING IF THIS IS GONNA WORK
		for(int i = 0; i < hpv.length; i++) {
			
		// This is all the processing required for a raw data sheet
		int size = Inputs.size_dilutionlst(raw_sheet);
		double[] data = new double[size];
		double[] dilution = Inputs.get_dilutions(raw_sheet, size); // gets the dilution list
		double[] ctrl = Inputs.ctrl_standards (ctrl_sheet, size);
		
		XSSFWorkbook output = new XSSFWorkbook();
		
		 
		XSSFSheet out_sheet = Outputs.create_sheet(output, hpv[i], dilution);
		Outputs.output_file(output);
		
		
		
		
		
		
		
		int df = Calculations.calculate_df(dilution); 
		
	    double[] log = new double[size];
		double[] logd = new double[size];
		double wPLL_slope;
		double pll_slope;
		double slope;
		double Xmean_REAL;
		double Ymean_REAL;

		// control and standards calculations should be here
		logd = Calculations.log_results(dilution);
		double ctrl_Ymean = Calculations.Ymean(logd);
		double ctrl_Xmean = Calculations.Xmean(logd);
		double SXX = Calculations.sxx(logd, ctrl_Xmean);
		double SXY = Calculations.sxy(logd, ctrl_Xmean, ctrl_Ymean);
		double first_slope = (SXY / SXX);
		double rfl_denominator = (ctrl_Xmean - (ctrl_Ymean / first_slope));
		double wPLL;
		double rfl;
		double pll;
		double correlation;
		double slope_ratio;
		
		int index = 1; //this is the index on the row of the output sheet 
		int pos = 0;
		int ending = (raw_sheet.getLastRowNum() - size);
		while (pos <= ending) { // counter for each line
			data = Inputs.line_raw(raw_sheet, "HPV 6", pos, size);

			// all the calculations go here
			log = Calculations.log_results(data);
			Xmean_REAL = Calculations.Xmean(log);
			Ymean_REAL = Calculations.Ymean(log);
			wPLL_slope = Calculations.slopewPLL(log, Xmean_REAL, Ymean_REAL, SXX, SXY);
			slope = Calculations.slope(log, Xmean_REAL, Ymean_REAL);
			pll_slope = ((first_slope + slope) / 2);
			
			// results
			wPLL = Calculations.wPLL(rf, df, wPLL_slope, Xmean_REAL, Ymean_REAL, ctrl_Xmean, ctrl_Ymean);
			rfl = Calculations.rfl(rf, df, rfl_denominator, first_slope, Xmean_REAL, Ymean_REAL);
			pll = Calculations.pll(rf, df, pll_slope, Xmean_REAL, Ymean_REAL, ctrl_Xmean, ctrl_Ymean);
			
			slope_ratio = (slope / first_slope);
			correlation = Calculations.correlation(data);
			
			
			//things I need to put. I AM MISSING THE RUN AND ID ONLY
			System.out.println(wPLL);
			System.out.println(rfl);
			System.out.println(pll);
			System.out.println(correlation);
			System.out.println(slope);
			System.out.println(slope_ratio);
			System.out.println(data);
			
			//Output.write_data(out_sheet, columns, style,index);
			pos = (pos + size); // counter used for extracting the next line
			index++;
		}
		}
	}
}
