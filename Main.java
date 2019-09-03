import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Locale;
import java.util.Scanner;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
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
		System.out.print("Path to the Excel file (should end in /FileName.xlsx): ");
		String input_file = in.nextLine();
		File input_path = new File(input_file);



		// loads the input excelfile
		FileInputStream fis = new FileInputStream(input_path);
		@SuppressWarnings("resource")
		XSSFWorkbook input = new XSSFWorkbook(fis);
		// XSSFWorkbook input = load_excel(input_path.getName());
		
		//all the parameters of the program will be read now
		XSSFSheet parameter_sheet = input.getSheet("Parameters");
		Inputs.read_parameters(parameter_sheet); 
		
		// Assigns names to the sheets
		XSSFSheet ctrl_sheet = input.getSheet(Inputs.standards);
		XSSFSheet rf_sheet = input.getSheet(Inputs.reference_factors);
		XSSFSheet cutoff_sheet = input.getSheet(Inputs.cut_off);


		//////////////////////////////////////////////////////////////////////////////////////////
		//////////////////////////////////////////////////////////////////////////////////////////
		//////////////////////////////////////////////////////////////////////////////////////////
		String[] hpv = Inputs.hpvlst(rf_sheet); // creates the list of all hpvs
		int max = Inputs.count_ints(rf_sheet); // counts how many different rf there is to do

		XSSFWorkbook output = new XSSFWorkbook();

		CellStyle error_cellstyle = Outputs.seronegative_cellstyle(output);
		CellStyle default_cellstyle = Outputs.default_cellstyle(output);
		CellStyle warning_cellstyle = Outputs.warning_cellstyle(output);

		//counter for all the raw data sheets
		for (int raw_counter = 0; raw_counter < Inputs.raw_data.length; raw_counter++) {
			XSSFSheet raw_sheet = input.getSheet(Inputs.raw_data[raw_counter]);
			
			// counter for viruses
			for (int master_counter = 0; master_counter < max; master_counter++) {

				// int [] rf = reference_factors (input,rf_sheet);
				double rf = Inputs.id_hpv(rf_sheet, hpv[master_counter]); // rf factor and id for that specific virus
				double cut_off_value = Inputs.cut_off(cutoff_sheet, hpv[master_counter]);

				// This is all the processing required for a raw data sheet
				int size = Inputs.size_dilutionlst(raw_sheet);
				double[] data = new double[size];
				double[] data_calculations = new double[size];

				double[] dilution = Inputs.get_dilutions(raw_sheet, size); // gets the dilution list
				double[] ctrl; // will be done in the while loop

				XSSFSheet out_sheet = Outputs.create_sheet(output, Inputs.raw_data[raw_counter], hpv[master_counter]);

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
				double[] fixed_data;

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

				System.out.println("processing");

				while (pos <= ending) { // counter for each line
					String[] run_id = Inputs.run_id(raw_sheet, pos);
					if (!run_id[0].equals(run_check)) {
						first_line = true;
						size = Inputs.size_dilutionlst(raw_sheet);
						data = new double[size];
						dilution = Inputs.get_dilutions(raw_sheet, size);
					}
					run_check = run_id[0];
					if (first_line) {
						ctrl = Inputs.ctrl_standards(ctrl_sheet, hpv[master_counter], size, run_id, dilution);

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
						data_results = Outputs.data_results(Inputs.id_dilution, ctrl, wPLL, rfl, pll, correlation,
								first_slope, slope_ratio);
						String temp = run_id[1];
						run_id[1] = Inputs.stand;
						Outputs.write_data(default_cellstyle, warning_cellstyle, out_sheet, index, run_id, data_results);
						run_id[1] = temp;
						index++;
						first_line = false;
					}

					data = Inputs.line_raw(raw_sheet, hpv[master_counter], pos, size);
					// data = Inputs.line_raw(raw_sheet, "HPV 6", pos, size);
					boolean seropositive = Inputs.seropositivity(cut_off_value, data);
					if (!seropositive) {
						Outputs.swrite_data(error_cellstyle, out_sheet, index, run_id, data);
//						Outputs.write_data(error_cellstyle, error_cellstyle, out_sheet, correlation_cut_off, slope_cut_off, sloperatio_cut_off, index, run_id, data);
						index++;
					}

					// only seropositive samples should be used for calculations
					if (seropositive) {
						// removing values to get the negative slope
						data_calculations = Calculations.fix_negative_slope(data);
						fixed_data = Calculations.fix_array(data_calculations);
						log = Calculations.log_results(fixed_data);
					  //log = Calculations.log_results(data);

						// calculations for lines other than the reference (first line)
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
						Outputs.sswrite_data(default_cellstyle, warning_cellstyle, out_sheet, index, run_id, data_results, data_calculations);
						index++;
					}
					pos = (pos + size); // counter used for extracting the next line
				}
				// inside for loop but outside while loop
			}
			// this is outside the virus loop
		
		}
		//this is outside the raw data loop
		Outputs.output_file(output, input); 
		System.out.println("file created");
	}
}
