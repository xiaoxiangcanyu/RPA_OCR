package HttpUtil;

import com.alibaba.fastjson.JSON;
import org.apache.commons.io.FileUtils;

import java.io.File;
import java.io.IOException;

public class Test {
    public static void main(String[] args) throws IOException {
        File file = new File("C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\0001.jpg");
        byte[] fileData = FileUtils.readFileToByteArray(file);//读取图片
        System.out.println(JSON.toJSON(fileData));
    }
}
