package generador.csv ;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.Properties;
import java.util.stream.Collectors;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

public class ExcelParser {
    public static void main(String[] args) {
        try {
            Properties prop = new Properties();
            loadProperties(prop);
            // Get the property values
            final String PATH_TO_WALK = prop.getProperty("path.to.walk");
            final String PATH_TO_WRITE = prop.getProperty("path.to.write");
            Boolean processed = false;

            List<File> filesInFolder = Files.walk(Paths.get(PATH_TO_WALK))
                    .filter(Files::isRegularFile)
                    .map(Path::toFile)
                    .collect(Collectors.toList());

            Workbook destinationWorkbook = new HSSFWorkbook();
            Sheet destinationSheet = destinationWorkbook.createSheet("ParsedData");
            int rowIndex = 0;

            for (File file : filesInFolder) {
                if (file.isFile() && file.getName().endsWith(".xlsx")) {
                    Workbook sourceWorkbook = WorkbookFactory.create(file);

                    for (Sheet sourceSheet : sourceWorkbook) {
                        for (Row sourceRow : sourceSheet) {
                            Cell telephoneCell = sourceRow.getCell(2); // Adjust column index for telephone number
                            Cell personCell = sourceRow.getCell(3); // Adjust column index for person name

                            if (telephoneCell != null && personCell != null) {
                                String telephone = telephoneCell.getStringCellValue();
                                String person = personCell.getStringCellValue();

                                String normalizedTelephone = normalizeTelephone(telephone);

                                // Check if the person already exists in the destination sheet
                                int existingRowIndex = findExistingPerson(destinationSheet, person);

                                if (existingRowIndex == -1) {
                                    // Person does not exist, create a new row
                                    Row destinationRow = destinationSheet.createRow(rowIndex++);
                                    Cell destinationTelephoneCell = destinationRow.createCell(6); // Column G
                                    Cell destinationPersonCell = destinationRow.createCell(2); // Column C

                                    destinationTelephoneCell.setCellValue(normalizedTelephone);
                                    destinationPersonCell.setCellValue(person);
                                } else {
                                    // Person exists, update the telephone number
                                    Row existingRow = destinationSheet.getRow(existingRowIndex);
                                    Cell existingTelephoneCell = existingRow.getCell(6); // Column G

                                    existingTelephoneCell.setCellValue(normalizedTelephone);
                                }

                                processed = true;
                            }
                        }
                    }

                    sourceWorkbook.close();
                }
            }

            if (processed) {
                FileOutputStream fos = new FileOutputStream(PATH_TO_WRITE);
                destinationWorkbook.write(fos);
                destinationWorkbook.close();
                fos.close();

                System.out.println("Parsing complete. Output file saved at: " + PATH_TO_WRITE);
            } else {
                System.out.println("Nothing to process.");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String normalizeTelephone(String telephone) {
        // Clean the telephone number by removing noise characters
        telephone = telephone.replaceAll("[^0-9]", "");

        // Normalize country code and area code if present
        if (telephone.startsWith("549")) {
            telephone = telephone.substring(3);
        } else if (telephone.startsWith("54")) {
            telephone = telephone.substring(2);
        } else if (telephone.startsWith("15")) {
            telephone = "261" + telephone.substring(2);
        }

        // Remove invisible spaces (Unicode character U+3161)
        telephone = telephone.replace("\u3161", "");

        return telephone;
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

    private static int findExistingPerson(Sheet sheet, String person) {
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            Cell personCell = row.getCell(2); // Column C

            if (personCell != null && personCell.getStringCellValue().equals(person)) {
                return i;
            }
        }
        return -1;
    }
}
