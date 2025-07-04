import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class ReadingPasswordProtectedExcelFile {
    public static void main(String[] args) {
        String excelFilePath = "./data/users.xlsx";
        String password = "karim";

        try (FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = WorkbookFactory.create(fis, password)) {

            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                for (Cell cell : row) {
                    System.out.printf("%-25s", getCellValue(cell));
                }
                System.out.println();
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) return "(null)";

        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> getFormulaCellValue(cell);
            case BLANK -> "";
            default -> "(UNSUPPORTED_TYPE)";
        };
    }

    private static String getFormulaCellValue(Cell cell) {
        return switch (cell.getCachedFormulaResultType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case BLANK -> "";
            default -> "(UNSUPPORTED_FORMULA_TYPE)";
        };
    }
}
