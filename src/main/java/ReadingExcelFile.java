import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class ReadingExcelFile {

    public static void main(String[] args) {
        String excelFilePath = "./data/users.xlsx";

        try (FileInputStream inputStream = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(inputStream)) {

            Sheet sheet = workbook.getSheetAt(0);
            int lastRow = sheet.getLastRowNum();
            int maxCols = getMaxColumns(sheet);

            for (int r = 0; r <= lastRow; r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                for (int c = 0; c < maxCols; c++) {
                    Cell cell = row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    System.out.printf("%-25s", getCellValue(cell));
                }
                System.out.println();
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static int getMaxColumns(Sheet sheet) {
        int maxCols = 0;
        for (Row row : sheet) {
            if (row != null && row.getLastCellNum() > maxCols) {
                maxCols = row.getLastCellNum();
            }
        }
        return maxCols;
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
