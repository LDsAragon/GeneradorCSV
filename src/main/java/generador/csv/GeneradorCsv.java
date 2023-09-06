package generador.csv;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.util.Iterator;
import java.util.Locale;
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

            // Determine the maximum column index
            int maxColumnIndex = 0;
            for (Row row : sheet) {
                int lastCellIndex = row.getLastCellNum();
                if (lastCellIndex > maxColumnIndex) {
                    maxColumnIndex = lastCellIndex;
                }
            }

            // Create a DecimalFormat to format numeric values as strings
            DecimalFormat decimalFormat = new DecimalFormat("0", DecimalFormatSymbols.getInstance(Locale.ENGLISH));
            decimalFormat.setMaximumFractionDigits(Integer.MAX_VALUE);

            // Iterate over rows
            for (Row row : sheet) {
                for (int columnIndex = 0; columnIndex <= maxColumnIndex; columnIndex++) {
                    Cell cell = row.getCell(columnIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

                    if (cell.getCellType() == CellType.NUMERIC) {
                        // Format numeric values as strings
                        String formattedValue = decimalFormat.format(cell.getNumericCellValue());
                        csvWriter.write(formattedValue);
                    } else {
                        // Write other cell types as is
                        csvWriter.write(cell.toString());
                    }

                    if (columnIndex < maxColumnIndex) {
                        // Add a comma between values (except for the last column)
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
            double durationInMilliseconds = (endTime - startTime) / 1000000;  //divide by 1000000 to get milliseconds.
            System.out.println("This process took " + durationInSeconds + " seconds");
            System.out.println("This process took " + durationInMilliseconds + " milliseconds");


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
