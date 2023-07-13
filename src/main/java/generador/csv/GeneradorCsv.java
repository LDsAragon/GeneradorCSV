package generador.csv;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.Properties;
import java.util.stream.Collectors;


public class GeneradorCsv {

    public static void main(String[] args) throws Exception {

        Properties prop = new Properties();
        loadProperties(prop);
        // get the property value and print it out
        final String PATH_TO_WALK = prop.getProperty("path.to.walk") ;

        try {

        List<File> filesInFolder = Files.walk(Paths.get(PATH_TO_WALK))
                .filter(Files::isRegularFile)
                .map(Path::toFile)
                .collect(Collectors.toList());

            System.out.println(Constants.NEW_LINE);
            System.out.println("FilesInFolder");
            System.out.println(filesInFolder);
            System.out.println("Size: "  + filesInFolder.size());

        }catch(Exception e){
            System.out.println(e);
            throw e ;
        }

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
