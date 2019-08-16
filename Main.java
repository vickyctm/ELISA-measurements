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
		System.out.println("Provide the name of the sheet that contains the value for seropositivity: ");
		String cut_off = in.nextLine();

		// ASK FOR THE FREAKING LIST OF DILUTIONS SUPER HELLA STUPID SHIT
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
		XSSFSheet cutoff_sheet = input.getSheet(cut_off);

		//////////////////////////////////////////////////////////////////////////////////////////
		//////////////////////////////////////////////////////////////////////////////////////////
		//////////////////////////////////////////////////////////////////////////////////////////
		String[] hpv = Inputs.hpvlst(rf_sheet); // creates the list of all hpvs
		int max = Inputs.count_ints(rf_sheet); // counts how many different rf there is to do

		XSSFWorkbook output = new XSSFWorkbook();

		CellStyle error_cellstyle = Outputs.seronegative_cellstyle(output);
		CellStyle default_cellstyle = Outputs.default_cellstyle(output);

		// for (int master_counter = 0; master_counter < max; master_counter++) {

		// int [] rf = reference_factors (input,rf_sheet); //extracts the rf needed for
		double rf = Inputs.id_hpv(rf_sheet, hpv); // rf factor and id for that specific virus
		double cut_off_value = Inputs.cut_off(cutoff_sheet, "HPV 6"); // hpv[master_counter]);

		// This is all the processing required for a raw data sheet
		int size = Inputs.size_dilutionlst(raw_sheet);
		double[] data = new double[size];
		double[] data_calculations = new double [size];
		
		double[] dilution = Inputs.get_dilutions(raw_sheet, size); // gets the dilution list
		double[] ctrl; // will be done in the while loop

		XSSFSheet out_sheet = Outputs.create_sheet(output, master, "HPV 6");// hpv[master_counter]);

		int df = Calculations.calculate_df(dilution);

		int index = 1; // this is the index on the row of the output sheet
		double[] log = new double[size];
		double[] log_ctrl = new double[size];
		double wPLL_slope;
		double pll_slope;
		double slope;
		double meanX;
		double meanY;
		double wPLL;
		double rfl;
		double pll;
		double correlation;
		double slope_ratio;
		double[] data_results;

		double ctrl_Ymean = 0;
		double ctrl_Xmean = 0;
		double SXX = 0;
		double SXY = 0;
		double first_slope = 0;
		double rfl_denominator = 0;
		// to check if this is a new run, if it is
		// get a new value from ctrls & standards
		String run_check = "0";
		boolean first_line = false;

		int pos = 0;
		int ending = (raw_sheet.getLastRowNum() - size);

		while (pos <= 500) {
			// while (pos <= ending) { // counter for each line

			String[] run_id = Inputs.run_id(raw_sheet, pos);
			if (!run_id[0].equals(run_check)) {
				first_line = true;
				size = Inputs.size_dilutionlst(raw_sheet);
				data = new double[size];
				dilution = Inputs.get_dilutions(raw_sheet, size);
			}
			run_check = run_id[0];
			if (first_line) {
				// ctrl = Inputs.ctrl_standards(ctrl_sheet, hpv[master_counter], size, run_id,
				// dilution);
				ctrl = Inputs.ctrl_standards(ctrl_sheet, "HPV 6", size, run_id, dilution);

				log_ctrl = Calculations.log_results(ctrl);
				ctrl_Ymean = Calculations.Ymean(log_ctrl);
				ctrl_Xmean = Calculations.Xmean(log_ctrl);
				SXX = Calculations.sxx(log_ctrl, ctrl_Xmean);
				SXY = Calculations.sxy(log_ctrl, ctrl_Xmean, ctrl_Ymean);
				first_slope = (SXY / SXX);
				rfl_denominator = (ctrl_Xmean - (ctrl_Ymean / first_slope));
				// first line
				wPLL_slope = Calculations.slopewPLL(log, ctrl_Xmean, ctrl_Ymean, SXX, SXY);
				wPLL = Calculations.wPLL(rf, df, wPLL_slope, ctrl_Xmean, ctrl_Ymean, ctrl_Xmean, ctrl_Ymean);
				rfl = Calculations.rfl(rf, df, rfl_denominator, first_slope, ctrl_Xmean, ctrl_Ymean);
				pll = Calculations.pll(rf, df, first_slope, ctrl_Xmean, ctrl_Ymean, ctrl_Xmean, ctrl_Ymean);
				slope_ratio = (first_slope / first_slope);
				correlation = Calculations.correlation(log_ctrl);
				data_results = Outputs.data_results(Inputs.id_dilution, ctrl, wPLL, rfl, pll, correlation, first_slope,
						slope_ratio);
				String temp = run_id[1];
				run_id[1] = Inputs.stand;
				Outputs.write_data(default_cellstyle, out_sheet, index, run_id, data_results);
				run_id[1] = temp;
				index++;
				first_line = false;
			}

			// data = Inputs.line_raw(raw_sheet, hpv[master_counter], pos, size);
			data = Inputs.line_raw(raw_sheet, "HPV 6", pos, size);
			boolean seropositive = Inputs.seropositivity(cut_off_value, data);
			if (!seropositive) {

				Outputs.write_data(error_cellstyle, out_sheet, index, run_id, data);
				System.out.println(index);
				index++;
			}

			// only seropositive samples should be used for calculations
			if (seropositive) {
				// removing values to get the negative slope
				//data_calculations = Calculations.fix_negative_slope(data);
				//log = Calculations.log_results(data_calculations);
				
				// calculations for lines other than the reference (first line)
				log = Calculations.log_results(data);
				meanX = Calculations.Xmean(log);
				meanY = Calculations.Ymean(log);
				wPLL_slope = Calculations.slopewPLL(log, meanX, meanY, SXX, SXY);
				slope = Calculations.slope(log, meanX, meanY);
				pll_slope = ((first_slope + slope) / 2);

				// to write
				wPLL = Calculations.wPLL(rf, df, wPLL_slope, meanX, meanY, ctrl_Xmean, ctrl_Ymean);
				wPLL = (wPLL / Inputs.factor);
				rfl = Calculations.rfl(rf, df, rfl_denominator, first_slope, meanX, meanY);
				rfl = (rfl / Inputs.factor);
				pll = Calculations.pll(rf, df, pll_slope, meanX, meanY, ctrl_Xmean, ctrl_Ymean);
				pll = (pll / Inputs.factor);
				slope_ratio = (slope / first_slope);
				correlation = Calculations.correlation(log);

				double d = dilution[0];
				data_results = Outputs.data_results(d, data, wPLL, rfl, pll, correlation, slope, slope_ratio);
				Outputs.write_data(default_cellstyle, out_sheet, index, run_id, data_results);
				System.out.println(index);
				index++;
			}
			pos = (pos + size); // counter used for extracting the next line
		}
		// Outputs.output_file(output); // for testing
		// }
		Outputs.output_file(output); // real place
	}
}
