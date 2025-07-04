import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class ReadingExcelFile {
    public static void main(String[] args) throws IOException {
        String excelFilePath = "./data/users.xlsx";

        try (FileInputStream inputstream = new FileInputStream(excelFilePath);
             XSSFWorkbook workbook = new XSSFWorkbook(inputstream)) {

            XSSFSheet sheet = workbook.getSheetAt(0);
            int rows = sheet.getLastRowNum();
            System.out.println("rows = " + rows);
            int cols = sheet.getRow(1).getLastCellNum();
            System.out.println("cols = " + cols);

            for (int r = 0; r <= rows; r++) {
                XSSFRow row = sheet.getRow(r);
                if (row == null) continue;

                for (int c = 0; c < cols; c++) {
                    XSSFCell cell = row.getCell(c);
                    System.out.printf("%-25s", getCellValue(cell));
                }
                System.out.println();
            }
        }
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                switch (cell.getCachedFormulaResultType()) {
                    case STRING:
                        return cell.getStringCellValue();
                    case NUMERIC:
                        return String.valueOf(cell.getNumericCellValue());
                    case BOOLEAN:
                        return String.valueOf(cell.getBooleanCellValue());
                    default:
                        return "";
                }
            default:
                return "";
        }
    }
}
