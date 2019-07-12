package HttpUtil;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.commons.io.FileUtils;
import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.ContentType;
import org.apache.http.entity.mime.HttpMultipartMode;
import org.apache.http.entity.mime.MultipartEntityBuilder;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.util.EntityUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 海关关税单OCR全图扫面，数据处理业务
 */
public class CustomsDuties {
    public static void main(String[] args) {
//        String filePath = args[0];
//        String fileName = args[1];
        String filePath = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AP\\俄罗斯工厂\\20104171332\\0001.jpg";
        String fileName = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AP\\俄罗斯工厂\\20104171332\\Excel_Value.xlsx";
        try {
            getData(filePath,fileName);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    public static void getData(String filePath,String fileName)throws Exception{
        File file = new File(filePath);
        byte[] fileData = FileUtils.readFileToByteArray(file);
        //OCR图片扫描
        String json = ocrImageFile(file.getName(), fileData);
        JSONObject jsonObject = JSON.parseObject(json);
        List<Map<String,String>> list = new ArrayList<>();//存储 结果数据
        //判断 返回的扫描结果
        if (jsonObject.get("result").toString().contains("success")) {
            JSONArray jsonArray = jsonObject.getJSONObject("ocrResult").getJSONArray("ranges");
            for (int i = 0; i < jsonArray.size(); i++) {
                Map<String ,String> map = new HashMap<>();
                JSONObject object = (JSONObject) jsonArray.get(i);
                JSONObject jsonObject2 = JSON.parseObject(object.get("range").toString());
                map.put("value",object.get("value").toString());
                map.put("x",jsonObject2.get("x").toString());
                map.put("width",jsonObject2.get("width").toString());
                map.put("y",jsonObject2.get("y").toString());
                map.put("height",jsonObject2.get("height").toString());
                list.add(map);
            }
        }
        if (list.size() > 0){
            excelOutput(list,fileName);
        }
    }

    /**
     * 导出数据表格
     * @param list
     * @param fileName
     * @throws Exception
     */
    public static void excelOutput(List<Map<String,String>> list,String fileName) throws Exception{
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        XSSFSheet xssfSheet = xssfWorkbook.createSheet("Sheet1");
        Row row0 = xssfSheet.createRow(0);
        String[] headers = new String[]{"Value","X","Width", "Y","Height"};
        for (int i = 0; i < headers.length; i++) {
            xssfSheet.setColumnWidth(i, 5000);
            XSSFCell cell = (XSSFCell) row0.createCell(i);
            XSSFRichTextString text = new XSSFRichTextString(headers[i]);
            cell.setCellValue(text);
        }
        int rowNum = 1;
        for (Map<String,String>  map : list){
            XSSFRow row = xssfSheet.createRow(rowNum);
            row.createCell(0).setCellValue(map.get("value"));
            row.createCell(1).setCellValue(map.get("x"));
            row.createCell(2).setCellValue(map.get("width"));
            row.createCell(3).setCellValue(map.get("y"));
            row.createCell(4).setCellValue(map.get("height"));
            rowNum++;
        }
        FileOutputStream fileOutputStream=new FileOutputStream(fileName);
        xssfWorkbook.write(fileOutputStream);

    }

    /**
     * OCR图片扫描
     * @param fileName
     * @param fileData
     * @return
     * @throws Exception
     */
    public static String ocrImageFile(String fileName,byte[] fileData) throws Exception{
         String baseUrl = "http://ocrserver.openserver.cn:8090/OcrServer/ocr/ocrImageByTemplate";
//        String baseUrl = "http://10.138.93.103:8080/OcrServer/ocr/directOcrImage";
        HttpPost post = new HttpPost(baseUrl);
        ContentType contentType = ContentType.create("multipart/form-data", Charset.forName("UTF-8"));
        MultipartEntityBuilder builder = MultipartEntityBuilder.create();
        builder.setMode(HttpMultipartMode.BROWSER_COMPATIBLE);
        builder.setCharset(Charset.forName("UTF-8"));
        builder.addBinaryBody("file", fileData, contentType, fileName);// 文件流
//        builder.addTextBody("imageType", imageType,contentType);// 类似浏览器表单提交，对应input的name和value
        HttpEntity entity = builder.build();
        post.setEntity(entity);
        CloseableHttpClient httpclient = HttpClients.createDefault();
        try {
            HttpResponse response = httpclient.execute(post);
            if(response.getStatusLine().getStatusCode() == 200){
                String result = EntityUtils.toString(response.getEntity(),"utf-8");
                return result;
            } else {
                throw new Exception(EntityUtils.toString(response.getEntity(),"utf-8"));
            }
        }finally {
            httpclient.close();
        }
    }
}
