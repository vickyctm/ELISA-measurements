import java.io.FileOutputStream;
import java.io.IOException;


import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
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
	
	//Creates a new sheet with the name of the virus. Adds the header.
	public static XSSFSheet create_sheet (XSSFWorkbook workbook, String type, double[] dilution) throws IOException{
		CreationHelper createHelper = workbook.getCreationHelper();
		XSSFSheet sheet = workbook.createSheet(type);
		
		Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.BLACK.getIndex());
        
        CellStyle header_style = workbook.createCellStyle();
        header_style.setFont(headerFont);
        header_style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        header_style.setFillPattern(FillPatternType.SOLID_FOREGROUND);;
        
		//header information
	     String[] header = new String[dilution.length + 8];
	     header[0] ="Run";
	     header[1] ="id";
	     for (int i = 2; i < (header.length - 6); i++) {
	         header[i] = String.valueOf(dilution[i-2]);
	     }
	     
	     header[header.length - 6] ="wPLL";
	     header[header.length - 5] ="rfl";
	     header[header.length - 4] ="PLL";
	     header[header.length - 3] ="correlation";
	     header[header.length - 2] ="slope";
	     header[header.length - 1] ="slope ratio";
	      
	     
	     Row row = sheet.createRow(0);
	        for(int i = 0; i < header.length; i++) {
	            Cell cell = row.createCell(i);
	            cell.setCellValue(header[i]);
	            cell.setCellStyle(header_style);
	            sheet.autoSizeColumn(i);
	        }
  
//	     try (FileOutputStream outputStream = new FileOutputStream("Program.xlsx")) {
//	            workbook.write(outputStream);
//	        }
	     return sheet;
	}
	
	public static double[] data_results (double[] data, double wPLL, double rfl, 
			double pll, double correlation, double slope, double slope_ratio) {
		
		double[] result = new double [data.length + 6]; 
		for(int i = 0; i < data.length; i++) {
			result[i] = data[i];
		}
		 result[data.length] = wPLL;
		 result[data.length + 1] = rfl;
		 result[data.length + 2] = pll;
		 result[data.length + 3] = correlation;
		 result[data.length + 4] = slope;
		 result[data.length + 5] = slope_ratio;
		
		return result;
	}
	
	//writes a row of data
    public static void write_data(XSSFSheet sheet, int index, String[] run_id, double[] data_results) {
    	Row row = sheet.createRow(index);
    	Cell cell = row.createCell(0);
    	cell.setCellValue(run_id[0]);
    	sheet.autoSizeColumn(0);
    	cell = row.createCell(1);
    	cell.setCellValue(run_id[1]);
    	sheet.autoSizeColumn(1);
    	
        for(int i = 2; i < (data_results.length + 2); i++) {
            cell = row.createCell(i);
            cell.setCellValue(data_results[i-2]);
//            cell.setCellStyle(style);
            sheet.autoSizeColumn(i);
        }
	}
    
    //writes the output to a file 
    public static void output_file(XSSFWorkbook workbook) throws IOException {
   	
    FileOutputStream file_out = new FileOutputStream("Program.xlsx");
    workbook.write(file_out);
    file_out.close();
    workbook.close();
	}
	
}
