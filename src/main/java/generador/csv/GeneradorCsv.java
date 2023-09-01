package generador.csv;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.Iterator;
import java.util.Properties;

public class GeneradorCsv {
    public static void main(String[] args) {
        long startTime = System.nanoTime();
        Properties prop = new Properties();
        loadProperties(prop);
        // Get the property values
        final String PATH_TO_WRITE = prop.getProperty("path.to.write");
        final String CSV_NAME = prop.getProperty("csv.name");
        final String EXCEL_NAME = prop.getProperty("excel.name");
        final String DELIMITER = prop.getProperty("csv.delimiter");

        try {

            FileInputStream excelFile = new FileInputStream(new File(PATH_TO_WRITE + EXCEL_NAME));
            Workbook workbook = new HSSFWorkbook(excelFile);

            // Create a CSV writer
            FileWriter csvFile = new FileWriter(PATH_TO_WRITE + CSV_NAME);
            BufferedWriter csvWriter = new BufferedWriter(csvFile);

            // Get the first sheet
            Sheet sheet = workbook.getSheetAt(0);

            // Iterate over rows and columns
            for (Row row : sheet) {
                Iterator<Cell> cellIterator = row.iterator();
                int currentColumn = 0; // Keep track of the current column index

                while (cellIterator.hasNext()) {
                    Cell currentCell = cellIterator.next();
                    int cellColumn = currentCell.getAddress().getColumn();

                    // Fill in empty columns with delimiters
                    while (currentColumn < cellColumn) {
                        csvWriter.write(DELIMITER);
                        currentColumn++;
                    }

                    // Write cell value to CSV
                    csvWriter.write(currentCell.toString());
                    currentColumn++;

                    // Add a comma for all cells except the last one in each row
                    if (cellIterator.hasNext()) {
                        csvWriter.write(DELIMITER);
                    }
                }
                // Add a new line character after each row
                csvWriter.newLine();
            }

            // Close the CSV writer and Excel file
            csvWriter.close();
            excelFile.close();

            System.out.println("Excel to CSV conversion completed.");

            long endTime = System.nanoTime();
            double durationInSeconds = (endTime - startTime) / 1000000000;  //divide by 1000000000 to get seconds.
            System.out.println("This process took " + durationInSeconds + " seconds");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void loadProperties(Properties prop) {
        // Load Properties Files
        try (InputStream input = new FileInputStream("src/main/resources/application.properties")) {
            // Load a properties file
            prop.load(input);
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

}
