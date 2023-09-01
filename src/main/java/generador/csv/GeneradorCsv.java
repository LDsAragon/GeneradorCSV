package generador.csv;

import generador.csv.enums.Cabecera;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;
import java.util.stream.Collectors;

public class GeneradorCsv {
    public static void main(String[] args) {
        long startTime = System.nanoTime();
        Properties prop = new Properties();
        loadProperties(prop);
        // Get the property values
        final String PATH_TO_WALK = prop.getProperty("path.to.walk");
        final String PATH_TO_WRITE = prop.getProperty("path.to.write");
        final String CSV_NAME = prop.getProperty("csv.name");
        final String EXCEL_NAME = prop.getProperty("excel.name");
        Boolean processed = false;
        int rowIndex = 1;

        try {

            FileInputStream excelFile = new FileInputStream(new File(PATH_TO_WRITE + EXCEL_NAME ));
            Workbook workbook = new HSSFWorkbook(excelFile);

            // Create a CSV writer
            FileWriter csvFile = new FileWriter( PATH_TO_WRITE + CSV_NAME);
            BufferedWriter csvWriter = new BufferedWriter(csvFile);

            // Get the first sheet
            Sheet sheet = workbook.getSheetAt(0);

            // Iterate over rows and columns
            Iterator<Row> iterator = sheet.iterator();
            while (iterator.hasNext()) {
                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();

                while (cellIterator.hasNext()) {
                    Cell currentCell = cellIterator.next();
                    // Write cell value to CSV
                    csvWriter.write(currentCell.toString());

                    // Add a comma between values (customize delimiter as needed)
                    csvWriter.write(",");
                }
                // Add a new line character after each row
                csvWriter.newLine();
            }

            // Close the CSV writer and Excel file
            csvWriter.close();
            excelFile.close();

            System.out.println("Excel to CSV conversion completed.");

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
