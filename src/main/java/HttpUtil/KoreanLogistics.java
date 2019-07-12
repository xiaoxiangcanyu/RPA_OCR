package HttpUtil;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;
import org.apache.commons.io.FileUtils;

import java.io.File;
import java.io.IOException;

import static DataClean.BaseUtil.ocrImageFile;

/**
 * 韩国物流
 */
public class KoreanLogistics {
    public static void main(String[] args) throws Exception {
        String picPath = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\一次物流-韩国-韩元图片\\0001.jpg";
        clearKoreanLogisticsFunction(picPath);
    }

    /**
     * OCR识别韩国物流的图片
     * @param picPath
     */
    private static void clearKoreanLogisticsFunction(String picPath) throws Exception {
        File file = new File(picPath);
        byte[] fileData = FileUtils.readFileToByteArray(file);//读取图片
        //OCR模板扫描 返回结果数据
        String json = ocrImageFile("PO_6430_KRW", file.getName(), fileData);
        System.out.println("输出韩国json:"+ json);
        JSONObject jsonObject = JSON.parseObject(json);
    }

}
