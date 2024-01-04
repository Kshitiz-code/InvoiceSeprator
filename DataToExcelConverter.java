import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class DataToExcelConverter {

    public static void main(String[] args) {
        String inputData = "QD DALUSXD .JAXBXCR 20231205200516 CAM S1NVOICE/CUS USA/SPL 83311/CPO 9122899/PNR 508390-2/INV 26657794/INQ 1/UNT EA/UNP 23507.40/IND 051223/INA 23507.40/ITC D/PAY 23507.40/ICR USD/SHT BB/BOL 704806560736/PSN 26657794";

        try {
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("DataSheet");

            Row headerRow = sheet.createRow(0);
            String[] headers = {"CUS", "SPL", "CPO", "PNR", "INV", "INQ", "PAY"};

            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
            }

            String[] data = extractData(inputData);

            Row dataRow = sheet.createRow(1);
            for (int i = 0; i < data.length; i++) {
                Cell cell = dataRow.createCell(i);
                cell.setCellValue(data[i]);
            }

            // Write the workbook to a file
            try (FileOutputStream fileOut = new FileOutputStream("output.xlsx")) {
                workbook.write(fileOut);
            }

            // Close the workbook
            workbook.close();
            System.out.println("Excel file created successfully!");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String[] extractData(String inputData) {
        String[] fields = {"CUS", "SPL", "CPO", "PNR", "INV", "INQ", "PAY"};
        String[] result = new String[fields.length];

        for (int i = 0; i < fields.length; i++) {
            String field = fields[i];
            String pattern = field + "\\s(\\S+)";
            String value = inputData.replaceAll(".*" + pattern + "\\s(\\S+).*", "$1").trim();
            result[i] = value;
        }
        
        for (int i = 0; i < result.length; i++) {
            String value = result[i];
            if (value.contains("/")) {
                result[i] = value.split("/")[0];
            }
        }



        return result;
    }

}
