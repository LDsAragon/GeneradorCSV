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
import java.util.List;
import java.util.Properties;
import java.util.stream.Collectors;

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
            generateHeader(destinationSheet);
            int rowIndex = 1;

            for (File file : filesInFolder) {
                if (file.isFile() && file.getName().endsWith(".xlsx")) {

                    System.out.println("Nombre del Archivo : " + file);

                    XSSFWorkbook sourceWorkbook = new XSSFWorkbook(file);

                    for (Sheet sourceSheet : sourceWorkbook) {

                        System.out.println("Nombre de la hoja : "  + sourceSheet.getSheetName());

                        for (int i = 10; i < sourceSheet.getPhysicalNumberOfRows(); i++) {

                            Row sourceRow = sourceSheet.getRow(i);

                            Cell telephoneCell = null ;
                            Cell personCell = null ;

                            if (sourceRow != null) {
                                telephoneCell = sourceRow.getCell(19); // Adjust column index for telephone number
                                personCell = sourceRow.getCell(2); // Adjust column index for person name
                            }

                            if (telephoneCell != null && personCell != null) {
                                String telephone = getCellPhoneValue(telephoneCell);
                                String person = personCell.getStringCellValue();

                                String normalizedTelephone = normalizeTelephone(telephone);

                                // Check if the person already exists in the destination sheet
                                int existingRowIndex = findExistingPerson(destinationSheet, person);

                                if ( !telephone.isBlank() && !person.isBlank() && existingRowIndex == -1) {
                                    // Person does not exist, create a new row
                                    Row destinationRow = destinationSheet.createRow(rowIndex++);
                                    Cell destinationTelephoneCell = destinationRow.createCell(Cabecera.C32.getPosicion()); // Column AG
                                    Cell destinationPersonCell = destinationRow.createCell(Cabecera.C1.getPosicion()); // Column A
                                    Cell destinationPersonCell2 = destinationRow.createCell(Cabecera.C2.getPosicion()); // Column A

                                    destinationTelephoneCell.setCellValue(normalizedTelephone);
                                    destinationPersonCell.setCellValue(person);
                                    destinationPersonCell2.setCellValue(person);


                                } else if (!telephone.isBlank() && !person.isBlank() ){
                                    // Person exists, update the telephone number
                                    Row existingRow = destinationSheet.getRow(existingRowIndex);
                                    Cell existingTelephoneCell = existingRow.getCell(32); // Column AG

                                    existingTelephoneCell.setCellValue(normalizedTelephone);
                                }
                                destinationSheet.autoSizeColumn(i);
                                processed = true;
                            }
                        }

                    }

                    sourceWorkbook.close();
                }
            }

            if (processed) {
                String rutaCompleta = PATH_TO_WRITE + "dummyExample.xlsx" ;
                FileOutputStream fos = new FileOutputStream(rutaCompleta);
                destinationWorkbook.write(fos);
                destinationWorkbook.close();
                fos.close();

                System.out.println("Parsing complete. Output file saved at: " + rutaCompleta);
            } else {
                System.out.println("Nothing to process.");
            }
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
            //throw new RuntimeException(e);
        }
    }

    private static String normalizeTelephone(String telephone) {
        // Clean the telephone number by removing noise characters
        telephone = telephone.replaceAll("[^0-9]", "");

        telephone = telephone.replaceAll("'", "");

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

    private static String getCellPhoneValue(Cell cell) {

        String result = "";

        if (cell.getCellType().equals(CellType.BLANK)) {
            result = "";
        }

        if (cell.getCellType().equals(CellType.STRING)) {
            result = cell.getStringCellValue();
        }

        if (cell.getCellType().equals(CellType.NUMERIC)) {
            result = String.valueOf(cell.getNumericCellValue());
        }

        if (cell.getCellType().equals(CellType.ERROR)) {
            result = String.valueOf(cell.getErrorCellValue());
        }

        if (cell.getCellType() != CellType.STRING
                && cell.getCellType() != CellType.NUMERIC
                && cell.getCellType() != CellType.ERROR
                && cell.getCellType() != CellType.BLANK
        ){
            System.out.println("CELL TYPE NOT KNOW: " + cell.getCellType() + " CELL data : " + cell );
        }

        return result;
    }

    private static void generateHeader(Sheet destinationSheet) {
        Row row = destinationSheet.createRow(0);
        Cabecera[] cabeceras = Cabecera.values();

        for (int i = 0; i < cabeceras.length ; i++) {
            row.createCell(i).setCellValue(cabeceras[i].getNombre());
            destinationSheet.autoSizeColumn(i);
        }

    }
}
