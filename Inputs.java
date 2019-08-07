import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
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

/**
 * @author Victoria Torres
 *
 */

public class Inputs {
	static String id = null;
	public static int call_counter = 0; // USED TO COUNT HOW MANY TIMES THE id_hpv METHOD WAS CALLED
	public static int rowindex = 0;

	// This method is used to load the excel files that would be used as inputs for
	// the program. File_name is provided by the user.
	public static XSSFWorkbook load_excel(String file_name)
			throws IOException, FileNotFoundException, InvalidFormatException {
	
		File excel_file = new File(file_name);
		FileInputStream fis = new FileInputStream(excel_file);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		
		// XSSFWorkbook workbook = new XSSFWorkbook(new File(file_name));

		return workbook;
	}

	// This method goes through the whole ref factors file and returns an array with
	// ALL the rf.
	public static int[] reference_factors(XSSFSheet rf_sheet) {
		int count = count_ints(rf_sheet);
		int[] rf_values = new int[count];
		int i = 0;
		int standard_row;
		int type_column;

		for (Row row : rf_sheet) {
			for (Cell cell : row) {
				if (cell.getCellType() == CellType.NUMERIC) {
					type_column = cell.getColumnIndex();
					standard_row = cell.getRowIndex();
					rf_values[i] = read_reference_factors(rf_sheet, standard_row, type_column);
					i++;
				}
			}
		}
		return rf_values;
	}

	// This method is called by reference_factors to obtain the value in the cell
	// (rf).
	public static int read_reference_factors(XSSFSheet rf_sheet, int standard_row, int type_column) {
		int rf = 0;

		XSSFCell rf_value = rf_sheet.getRow(standard_row).getCell(type_column);
		rf = (int) rf_value.getNumericCellValue();

		return rf;
	}

	// gets the number of cells in a document that have ints inside
	// used for reference_factors
	public static int count_ints(XSSFSheet rf_sheet) {
		int count = 0;
		for (Row row : rf_sheet) {
			for (Cell cell : row) {
				if (cell.getCellType() == CellType.NUMERIC) {
					count++;
				}
			}
		}
		return count;
	}

	// method used to find the position where the standard is located
	public static int find_row(XSSFSheet sheet, String cell_content) {
		for (Row row : sheet) {
			for (Cell cell : row) {
				if (cell.getCellType() == CellType.STRING) {
					if (cell.getRichStringCellValue().getString().trim().equals(cell_content)) {
						return row.getRowNum();
					}
				}
			}
		}
		return 0;
	}

	// methods used to find the position where the type is located
	public static int find_column(XSSFSheet sheet, String cell_content) {
		for (Row row : sheet) {
			for (Cell cell : row) {
				if (cell.getCellType() == CellType.STRING) {
					if (cell.getRichStringCellValue().getString().trim().equals(cell_content)) {
						return cell.getColumnIndex();
					}
				}
			}
		}
		return 0;
	}

	// returns a list with all the names of HPVs present in the reference factors
	// sheet
	public static String[] hpvlst(XSSFSheet rf_sheet) {
		Row hpvRow = rf_sheet.getRow(0);
		int numCol = hpvRow.getLastCellNum();
		String[] hpv = new String[numCol - 1];
		String content;

		// gets the name of all the HPVs
		for (int j = 1; j <= (numCol - 1); j++) {
			content = hpvRow.getCell(j).getStringCellValue();
			hpv[j - 1] = content;
		}
		return hpv;
	}

	// The HPV list is used as a sort of counter.
	public static double id_hpv(XSSFSheet rf_sheet, String[] hpv) {
		// int master_counter = count_ints (rf_sheet);
		call_counter++;
		double rf = 0;

		// THIS IS HOW YOU ITERATE THROUGH A COLUMN!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
		int colindex = find_column(rf_sheet, hpv[call_counter - 1]);
		
		
		for (rowindex = 0; rowindex <= rf_sheet.getLastRowNum(); rowindex++) {
			Row row = rf_sheet.getRow(rowindex);
			Cell cell = row.getCell(colindex);
			if (cell.getCellType() == CellType.NUMERIC) {
				rf = cell.getNumericCellValue();
				id = row.getCell(0).getStringCellValue();
				// If you put a break here you can get the first value
				// do that and save the position so next time it is called
				// you start from that row
				// a while loop might be better than a for loops
			}
		}

		// int doesitwork = (call_counter-1);
		// int colindex = find_column(rf_sheet,hpv[(call_counter-1)]);
		// int check;
		//
		// for(int rowindex = 0; rowindex <= rf_sheet.getLastRowNum(); rowindex++) {
		// Row row = rf_sheet.getRow(rowindex);
		// Cell cell = row.getCell(colindex);
		// if(cell.getCellType() == CellType.NUMERIC) {
		// rf = (int) cell.getNumericCellValue();
		// id = row.getCell(0).getStringCellValue();
		//
		// check = (rowindex + 1);
		// Row checkrow = rf_sheet.getRow(check);
		// Cell checkcell = checkrow.getCell(colindex);
		// if(checkcell.getCellType() == CellType.NUMERIC) {
		// call_counter--;
		// }
		// break;
		// }
		// }

		return rf;
	}

	// This method checks the number of dilutions there is by comparing the id and
	// run values
	public static int size_dilutionlst(XSSFSheet raw_sheet) {
		// column indexes
		int RUN = 0;
		int ID = 1;

		int size = 1;

		// Initial cells used to then compared with the others
		Row ROW1 = raw_sheet.getRow(2);
		Cell R1 = ROW1.getCell(RUN);
		Cell ID1 = ROW1.getCell(ID);

		for (int rowindex = 2; rowindex <= 10; rowindex++) {
			Row row = raw_sheet.getRow(rowindex);
			Cell r_cell = row.getCell(RUN);
			Cell id_cell = row.getCell(ID);
			if (((r_cell.getNumericCellValue()) == (R1.getNumericCellValue()))
					&& (id_cell.getStringCellValue().equals(ID1.getStringCellValue()))) {
				size++;
			}
		}
		return size;
	}

	// Gets the values of dilutions
	public static double[] get_dilutions(XSSFSheet raw_sheet, int size) {
		int DILUTION = 2;
		double[] dilution = new double[size];

		for (int rowindex = 1; rowindex <= size; rowindex++) {
			Row row = raw_sheet.getRow(rowindex);
			Cell d_cell = row.getCell(DILUTION);
			dilution[(rowindex - 1)] = d_cell.getNumericCellValue();
		}
		return dilution;
	}

	// Extracts a line of the raw data document
	public static double[] line_raw(XSSFSheet raw_sheet, String type, int pos, int size) {
		int type_col = find_column(raw_sheet, type);
		double[] data = new double[size];

		for (int rowindex = (pos + 1); rowindex < (size + (pos + 1)); rowindex++) {
			Row row = raw_sheet.getRow(rowindex);
			Cell data_cell = row.getCell(type_col);
			if (data_cell.getCellType() == CellType.NUMERIC) {
				data[((rowindex - (pos + 1)))] = data_cell.getNumericCellValue();
			} else
				data[(rowindex - (pos + 1))] = 0;
		}
		return data;
	}
	
	public static double[] ctrl_standards (XSSFSheet sheet, int size) {
		double[] ctrl= new double[size];
		
		
		
		
		return ctrl;
	}
	
}
