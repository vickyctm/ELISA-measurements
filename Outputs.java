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
        
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);
        headerCellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);;
        
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
	      
	     write_data(sheet, header,headerCellStyle, 0);
  
//	     try (FileOutputStream outputStream = new FileOutputStream("Program.xlsx")) {
//	            workbook.write(outputStream);
//	        }
	     return sheet;
	}
	
	//writes a row of data
    public static void write_data(XSSFSheet sheet, String[] columns,CellStyle style, int index) {
    	Row row = sheet.createRow(index);
        for(int i = 0; i < columns.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(style);
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
