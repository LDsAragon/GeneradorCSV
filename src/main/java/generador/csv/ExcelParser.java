package generador.csv;

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
            // get the property value and print it out
            final String PATH_TO_WALK = prop.getProperty("path.to.walk") ;
            final String PATH_TO_WRITE = prop.getProperty("path.to.write") ;
            Boolean processed = false ;


            List<File> filesInFolder = Files.walk(Paths.get(PATH_TO_WALK))
                    .filter(Files::isRegularFile)
                    .map(Path::toFile)
                    .collect(Collectors.toList());

            Workbook destinationWorkbook = new HSSFWorkbook();
            Sheet destinationSheet = destinationWorkbook.createSheet("ParsedData");
            int rowIndex = 0;

            System.out.println("============= FILES IN PATH =============");
            for (File file : filesInFolder) {

                System.out.println(file.toString());

                if (file.isFile() && file.getName().endsWith(".xlsx")) {

                    Workbook sourceWorkbook = WorkbookFactory.create(file);

                    for (Sheet sourceSheet : sourceWorkbook) {
                        for (Row sourceRow : sourceSheet) {
                            Cell telephoneCell = sourceRow.getCell(0);
                            Cell personCell = sourceRow.getCell(1);

                            if (telephoneCell != null && personCell != null) {
                                String telephone = telephoneCell.getStringCellValue();
                                String person = personCell.getStringCellValue();

                                String normalizedTelephone = normalizeTelephone(telephone);

                                Row destinationRow = destinationSheet.createRow(rowIndex++);
                                Cell destinationTelephoneCell = destinationRow.createCell(0);
                                Cell destinationPersonCell = destinationRow.createCell(1);

                                destinationTelephoneCell.setCellValue(normalizedTelephone);
                                destinationPersonCell.setCellValue(person);
                                processed = true ;
                            }
                        }
                    }

                    sourceWorkbook.close();
                }
            }

            if (processed) {

                // Create the destination folder if it doesn't exist
                File destinationFolder = new File(PATH_TO_WRITE).getParentFile();
                if (!destinationFolder.exists()) {
                    destinationFolder.mkdirs();
                }


                FileOutputStream fos = new FileOutputStream(PATH_TO_WRITE);
                destinationWorkbook.write(fos);
                destinationWorkbook.close();
                fos.close();

                System.out.println("Parsing complete. Output file saved at: " + PATH_TO_WRITE);
            } else {
                System.out.println("============= NOTHING DONE =============");
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
        // Load Properties Files !
        try (InputStream input = new FileInputStream("src/main/resources/application.properties")) {

            // load a properties file
            prop.load(input);

        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }
}
