import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class ReadingExcelFileWithIterator {

    public static void main(String[] args) {
        String excelFilePath = "./data/users.xlsx";

        try (FileInputStream inputStream = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(inputStream)) {
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
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> getFormulaCellValue(cell);
            default -> "";
        };
    }

    private static String getFormulaCellValue(Cell cell) {
        return switch (cell.getCachedFormulaResultType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            default -> "";
        };
    }
}
