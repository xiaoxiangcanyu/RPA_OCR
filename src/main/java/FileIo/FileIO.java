package FileIo;

import java.io.File;
import java.io.IOException;
import java.util.Date;

public class FileIO {
    public static void main(String args[]) throws IOException {
        FileIO fileIO=new FileIO();
      System.out.println(new File(fileIO.getClass().getResource("/").getPath()).getParent());


    }
}
