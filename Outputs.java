import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author Victoria Torres
 *
 */

public class Outputs {

	// Creates a new sheet with the name of the virus. Adds the header.
	public static XSSFSheet create_sheet(XSSFWorkbook workbook, String month, String type) throws IOException {
		String name = (month + "  " + type + "  " + Inputs.stand);
		XSSFSheet sheet = workbook.createSheet(name);

		Font headerFont = workbook.createFont();
		headerFont.setBold(true);
		headerFont.setFontHeightInPoints((short) 14);
		headerFont.setColor(IndexedColors.BLACK.getIndex());

		CellStyle header_style = workbook.createCellStyle();
		header_style.setFont(headerFont);
		header_style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
		header_style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		header_style.setBorderTop(BorderStyle.THIN);
		header_style.setBorderBottom(BorderStyle.THIN);
		header_style.setBorderLeft(BorderStyle.THIN);
		header_style.setBorderRight(BorderStyle.THIN);

		// header information
		String[] header = { "Run", "id", "50", "150", "450", "1350", "4050", "12150", "36450", "109350", "328050",
				"984150", "wPLL", "rfl", "PLL", "correlation", "slope", "slope ratio" };

		Row row = sheet.createRow(0);
		for (int i = 0; i < header.length; i++) {
			Cell cell = row.createCell(i);
			cell.setCellValue(header[i]);
			cell.setCellStyle(header_style);
			sheet.autoSizeColumn(i);
		}
		return sheet;
	}

	// puts the data and results in the same array
	public static double[] data_results(double dilution, double[] data, double wPLL, double rfl, double pll,
			double correlation, double slope, double slope_ratio) {
		int num_of_dilutions = data.length;
		double[] result = new double[16]; // everything except the run and id
		int start = 0;
		int index = 0;
		String[] header = {"Run", "id", "50.0", "150.0", "450.0", "1350.0", "4050.0", "12150.0", "36450.0", 
			"109350.0", "328050.0", "984150.0", "wPLL", "rfl", "PLL", "correlation", "slope", "slope ratio" };

		if (num_of_dilutions != 10 && dilution != 50) {
			// int size = (10 - num_of_dilutions);
			// result = new double[num_of_dilutions + size + 6];
			for (int i = 2; i < 12; i++) {
				if (Double.toString(dilution).equals(header[i])) {
					for (start = (i - 2); start <= num_of_dilutions; start++) {
						result[start] = data[index];
						index++;
						// break;
					}
				}
			}
		} else {
			// result = new double[num_of_dilutions + 6];
			for (int i = 0; i < data.length; i++) {
				result[i] = data[i];
			}
		}

		result[result.length - 6] = wPLL;
		result[result.length - 5] = rfl;
		result[result.length - 4] = pll;
		result[result.length - 3] = correlation;
		result[result.length - 2] = slope;
		result[result.length - 1] = slope_ratio;

		return result;
	}

	// writes a row of data
	public static void write_data(CellStyle style, XSSFSheet sheet, int index, String[] run_id, double[] data_results) {
		Row row = sheet.createRow(index);
		Cell cell;
		for (int i = 0; i < run_id.length; i++) {
			cell = row.createCell(i);
			cell.setCellValue(run_id[i]);
			cell.setCellStyle(style);
			sheet.autoSizeColumn(i);
		}

		int size = (data_results.length + 2);

		for (int i = 2; i < size; i++) {
			cell = row.createCell(i);
			if (data_results[i - 2] == 0) {
				// this sets cells to blank
			}else {
				cell.setCellValue(data_results[i - 2]);
				cell.setCellStyle(style);
			}
			sheet.autoSizeColumn(i);
		}
	}

	public static CellStyle seronegative_cellstyle(XSSFWorkbook workbook) {
		CellStyle style = workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);

		return style;
	}

	public static CellStyle default_cellstyle(XSSFWorkbook workbook) {
		CellStyle style = workbook.createCellStyle();
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		return style;
	}

	// writes the output to a file
	public static void output_file(XSSFWorkbook workbook) throws IOException {

		FileOutputStream file_out = new FileOutputStream("Serology.xlsx");
		workbook.write(file_out);
		file_out.close();
		workbook.close();
	}

}
