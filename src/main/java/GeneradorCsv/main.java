package GeneradorCsv;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class main {

    public static void main(String[] args) throws Exception {


        final String PATH_TO_WALK = "E:\\Archivos\\Escritorio\\Misc" ;

        try {
                Stream<Path> paths = Files.walk(Paths.get(PATH_TO_WALK)) ;
            paths
                    .filter(Files::isRegularFile)
                    .forEach(System.out::println);



        List<File> filesInFolder = Files.walk(Paths.get(PATH_TO_WALK))
                .filter(Files::isRegularFile)
                .map(Path::toFile)
                .collect(Collectors.toList());

            System.out.println("filesInFolder");
            System.out.println(filesInFolder);
            System.out.println("Size:"  + filesInFolder.size());

        }catch(Exception e){
            System.out.println(e);
            throw e ;
        }

    }

}
