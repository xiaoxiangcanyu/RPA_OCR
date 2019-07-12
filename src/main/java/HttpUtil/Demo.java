package HttpUtil;

import DataClean.BaseUtil;
import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Demo extends BaseUtil {
    public static void main(String[] args) {
//        String filePath = args[0];
//        String fileName = args[1];
        String filePath = "F:\\img\\泰国测试样本\\0123\\泰国工厂AP\\6560-FAC-1900957142.jpg";
        String fileName = "F:\\img\\泰国测试样本\\0123\\泰国工厂AP\\Excel.xlsx";
        try {
            getData(filePath,fileName);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 图片扫描，返回数据处理
     * @param filePath
     * @param fileName
     * @throws Exception
     */
    public static void getData(String filePath, String fileName) throws Exception{
        File file = new File(filePath);
        byte[] fileData = FileUtils.readFileToByteArray(file);
        //模板扫描
        String json = ocrImageFile("APTha_0928", file.getName(), fileData);
        System.out.println("json" + json);
        JSONObject jsonObject = JSON.parseObject(json);
        List<Map<String,String>> list = new ArrayList<>();
        //判断返回数据
        if (jsonObject.get("result").toString().contains("success")) {
            Map<String,String> map = new HashMap<>();
            //
            JSONArray jsonArray = jsonObject.getJSONObject("ocrResult").getJSONArray("ranges");
            for (int i = 0; i < jsonArray.size(); i++) {
                JSONObject object = (JSONObject) jsonArray.get(i);
                String name = object.get("rangeId").toString();
                if (!"item".equals(name)){
                    map.put(object.get("rangeId").toString(),object.get("value").toString());
                }
            }
            //获取扫描表格数据
            JSONArray jsonArrayRowDatas = jsonArray.getJSONObject(2).getJSONArray("rowDatas");
            for (int i = 0; i < jsonArrayRowDatas.size(); i++) {
                JSONArray jsonArray1 = jsonArrayRowDatas.getJSONArray(i);
                Map<String,String> map1 = new HashMap<>();
                for (int j = 0; j < jsonArray1.size(); j++) {
                    JSONObject object = (JSONObject) jsonArray1.get(j);
                    map1.put(object.get("columnId").toString(),object.get("value").toString());
                }
                list.add(map1);
            }
            for (Map<String,String>  m : list){
                for (String key : map.keySet()) {
                    m.put(key,map.get(key));
                }
            }
        }
        if (list.size() > 0){
            excelOutput(list,fileName);
        }
    }

    /**
     * 导出 数据表格
     * @param list
     * @param fileName
     * @throws Exception
     */
    public static void excelOutput( List<Map<String,String>> list,String fileName) throws Exception{
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        XSSFSheet xssfSheet = xssfWorkbook.createSheet("Sheet1");
        Row row0 = xssfSheet.createRow(0);
        Map<String,String> map = list.get(0);
        String[] headers = new String[map.size()];
        int index = 0;
        for (String key : map.keySet()) {//查看
            headers[index] = key;
            index++;
        }
        for (int i = 0; i < headers.length; i++) {
            XSSFCell cell = (XSSFCell) row0.createCell(i);
            XSSFRichTextString text = new XSSFRichTextString(headers[i]);
            cell.setCellValue(text);
            xssfSheet.setColumnWidth(i, 5000);
        }
        int rowNum = 1;
        for (Map<String,String>  m : list){
            XSSFRow row = xssfSheet.createRow(rowNum);
            for (String key : m.keySet()) {
                for (int i = 0; i < headers.length; i++) {
                    String name = headers[i];
                    if (key.equals(name)){
                        row.createCell(i).setCellValue(m.get(key));
                    }
                }
            }
            rowNum++;
        }
        FileOutputStream fileOutputStream=new FileOutputStream(fileName);
        xssfWorkbook.write(fileOutputStream);
    }
}
