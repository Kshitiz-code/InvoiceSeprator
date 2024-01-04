import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class invToexcel {

    public static void main(String[] args) {
        // Specify the path to the folder containing input files
        String folderPath = "C:\\Users\\KSGUPTA\\Documents\\INVOICES";

        // Specify the output Excel file
        String outputFilePath = "C:\\Users\\KSGUPTA\\Documents\\Invoice.xlsx";

        try {
            // Get all files in the folder
            File folder = new File(folderPath);
            File[] files = folder.listFiles();

            if (files != null) {
                Workbook workbook = new XSSFWorkbook();
                Sheet sheet = workbook.createSheet("Separated Invoice Details");

                Row headerRow = sheet.createRow(0);
                String[] headers = {"CUS", "SPL", "CPO", "PNR", "INV", "INQ", "PAY"};

                for (int i = 0; i < headers.length; i++) {
                    Cell cell = headerRow.createCell(i);
                    cell.setCellValue(headers[i]);
                }

                for (File inputFile : files) {
                    // Read input data from each file
                    String inputData = readInputFromFile(inputFile);

                    // Create a new data row for each file
                    Row dataRow = sheet.createRow(sheet.getLastRowNum() + 1);

                    String[] data = extractData(inputData);

                    for (int i = 0; i < data.length; i++) {
                        Cell cell = dataRow.createCell(i);
                        cell.setCellValue(data[i]);
                    }
                }

                // Write the workbook to a file
                try (FileOutputStream fileOut = new FileOutputStream(outputFilePath)) {
                    workbook.write(fileOut);
                }

                // Close the workbook
                workbook.close();
                System.out.println("Invoice file created successfully!");

            } else {
                System.out.println("No files found in the specified folder.");
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String readInputFromFile(File file) throws IOException {
        // Read content from the file
        StringBuilder content = new StringBuilder();
        try (BufferedReader reader = new BufferedReader(new FileReader(file))) {
            String line;
            while ((line = reader.readLine()) != null) {
                content.append(line);
            }
        }
        return content.toString();
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
