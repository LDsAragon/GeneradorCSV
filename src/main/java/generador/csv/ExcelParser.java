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
            long startTime = System.nanoTime();
            Properties prop = new Properties();
            loadProperties(prop);
            // Get the property values
            final String PATH_TO_WALK = prop.getProperty("path.to.walk");
            final String PATH_TO_WRITE = prop.getProperty("path.to.write");
            final String EXCEL_NAME = prop.getProperty("excel.name");
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

                    //System.out.println("Nombre del Archivo : " + file);

                    XSSFWorkbook sourceWorkbook = new XSSFWorkbook(file);

                    for (Sheet sourceSheet : sourceWorkbook) {

                        System.out.println("Nombre de la hoja : " + sourceSheet.getSheetName() + " Nombre del Archivo : " + file);

                        for (int i = 10; i < sourceSheet.getPhysicalNumberOfRows(); i++) {

                            //System.out.println("FIla: " + i); //DEBUG
                            Row sourceRow = sourceSheet.getRow(i);

                            Cell telephoneCell = null;
                            Cell personCell = null;

                            if (sourceRow != null) {
                                telephoneCell = sourceRow.getCell(Constants.CELDAS_TELEFONOS_1); // Adjust column index for telephone number

                                if (telephoneCell == null) {
                                    telephoneCell = sourceRow.getCell(Constants.CELDAS_TELEFONOS_2);
                                }

                                personCell = sourceRow.getCell(Constants.CELDAS_NAME); // Adjust column index for person name
                            }

                            if (telephoneCell != null && personCell != null) {
                                String telephone = getCellPhoneValue(telephoneCell);

                                /* If there is no string value in person cell empty it */
                                String person;
                                try {
                                    person = personCell.getStringCellValue();
                                } catch (IllegalStateException ilegaIllegalStateException) {
                                    person = Constants.VACIO;
                                }

                                String normalizedTelephone = normalizeTelephone(telephone);

                                // Check if the person already exists in the destination sheet

                                if (!normalizedTelephone.isBlank() && !person.isBlank()) {
                                    // Person does not exist, create a new row
                                    Row destinationRow = destinationSheet.createRow(rowIndex++);
                                    Cell destinationTelephoneCell = destinationRow.createCell(Cabecera.C32.getPosicion()); // Column AG
                                    Cell destinationPersonCell = destinationRow.createCell(Cabecera.C1.getPosicion()); // Column A
                                    Cell destinationPersonCell2 = destinationRow.createCell(Cabecera.C2.getPosicion()); // Column A

                                    destinationTelephoneCell.setCellValue(Long.parseLong(normalizedTelephone));
                                    destinationPersonCell.setCellValue(person);
                                    destinationPersonCell2.setCellValue(person);


                                }
                                processed = true;
                            }
                        }

                    }

                    sourceWorkbook.close();
                }
            }

            if (processed) {

                destinationSheet.autoSizeColumn(Cabecera.C1.getPosicion());
                destinationSheet.autoSizeColumn(Cabecera.C2.getPosicion());
                destinationSheet.autoSizeColumn(Cabecera.C32.getPosicion());

                String finalRoute = PATH_TO_WRITE + EXCEL_NAME;
                FileOutputStream fos = new FileOutputStream(finalRoute);
                destinationWorkbook.write(fos);
                destinationWorkbook.close();
                fos.close();

                System.out.println("Parsing complete. Output file saved at: " + finalRoute);
            } else {
                System.out.println("Nothing to process.");
            }

            long endTime = System.nanoTime();
            double durationInMilliseconds = (endTime - startTime) / 1000000;  //divide by 1000000 to get milliseconds.
            double durationInSeconds = (endTime - startTime) / 1000000000;  //divide by 1000000000 to get seconds.
            System.out.println("This process took " + durationInSeconds + " seconds");

        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    private static String normalizeTelephone(String telephone) {
        // Clean the telephone number by removing noise characters
        telephone = telephone.replaceAll("E9", "");
        telephone = telephone.replaceAll("[^0-9]", "");
        telephone = telephone.replaceAll("\"", "");
        telephone = telephone.replaceAll("'", "");
        // Remove invisible spaces (Unicode character U+3161)
        telephone = telephone.replace("\u3161", "");

        // Filter BUBATEC phone
        if (telephone.equals("2617198290")) {
            telephone = Constants.VACIO;
        }

        // Normalize country code and area code if present
        if (telephone.startsWith(Constants.STR_549)) {
            telephone = telephone.substring(3);
        } else if (telephone.startsWith(Constants.STR_54)) {
            telephone = telephone.substring(2);
        } else if (telephone.startsWith(Constants.STR_15)) {
            telephone = Constants.STR_261 + telephone.substring(2);
        }
        if (telephone.startsWith(Constants.STR_0)) {
            telephone = telephone.substring(1);
        }

        // Remove strange or unusual phones
        if (!telephone.isBlank() && telephone.length() != 10) {
            telephone = Constants.VACIO;
        }

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

    private static String getCellPhoneValue(Cell cell) {

        String result = Constants.VACIO;

        if (cell.getCellType().equals(CellType.BLANK)) {
            result = Constants.VACIO;
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
                && cell.getCellType() != CellType.BLANK) {
            // System.out.println("CELL TYPE NOT KNOW: " + cell.getCellType() + " CELL data : " + cell ); //DEBUG
        }

        return result;
    }

    private static void generateHeader(Sheet destinationSheet) {
        Row row = destinationSheet.createRow(0);
        Cabecera[] header = Cabecera.values();

        for (int i = 0; i < header.length; i++) {
            row.createCell(i).setCellValue(header[i].getNombre());
        }

    }
}
