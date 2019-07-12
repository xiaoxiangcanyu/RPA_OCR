package HttpUtil;

import DataClean.BaseUtil;
import DataClean.RussiaGTD;
import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.util.regex.Pattern;
/**
 * 俄罗斯海关单数据处理
 */
public class Russia_GTD extends BaseUtil {
    public static void main(String[] args) {
//        String companyCode = args[0];
//        String filePath = args[1];
//        String excel_filePath = args[2];
        String companyCode = "62F0-HQ";
        String filePath= "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\俄罗斯海关单和发票\\1\\GTD-0.jpg";
        String excel_filePath= "F:\\test\\GTD\\GTDxls.xlsx";
        String fileName = filePath.substring(filePath.lastIndexOf("\\") + 1);//截取图片路径
        try {
            clearGTDFunction(companyCode, filePath, excel_filePath);
        } catch (Exception e) {
            String msg = "发票数据扫描不完整";
            templateNot_GTD(companyCode,fileName,excel_filePath, msg);
        }
    }
    public static void clearGTDFunction(String companyCode, String filePath, String excel_filePath) throws Exception{
        String gtdNumber = "";
        String gtdQuantity = "";
        String gtdAmount = "";
        String fileName = filePath.substring(filePath.lastIndexOf("\\") + 1);//截取图片路径
        File file = new File(filePath);
        byte[] fileData = FileUtils.readFileToByteArray(file);//读取图片
        String json = ocrImageFile("APSigRusGTD", file.getName(), fileData);
        System.out.println("输出OCR扫描json:"+ json);
        JSONObject jsonObject = JSON.parseObject(json);
        if (jsonObject.get("result").toString().contains("success")) {
            //处理json数据
            JSONArray jsonArray = jsonObject.getJSONObject("ocrResult").getJSONArray("ranges");
            for (int i = 0; i <jsonArray.size(); i++) {
                JSONObject object = (JSONObject) jsonArray.get(i);
                switch (object.get("rangeId").toString()){
                    case "GTDNumber":
                        if (object.get("value") != null){
                            gtdNumber = object.get("value").toString();
                        }else {
                            gtdNumber = "";
                        }
                        break;
                    case "GTDQuantity":
                        if (object.get("value") != null){
                            gtdQuantity = checkQuantity(object.get("value").toString());
                        }else {
                            gtdQuantity = "";
                        }
                        break;
                    case "GTDAmount":
                        if (object.get("value") != null){
                            gtdAmount = convertData(object.get("value").toString());
                        }else {
                            gtdAmount = "";
                        }
                        break;
                }
            }
            RussiaGTD russiaGTD = new RussiaGTD();
            russiaGTD.setCompanyCode(companyCode);
            russiaGTD.setFileName(fileName);
            String status = "";
            if (!"".equals(gtdNumber) && !"".equals(gtdQuantity) && !"".equals(gtdAmount)){
                russiaGTD.setGtdNumber(gtdNumber);
                russiaGTD.setGtdQuantity(gtdQuantity);
                russiaGTD.setGtdAmount(gtdAmount);
                russiaGTD.setStatus("OK");
            }else {
                if ("".equals(gtdNumber)){
                    status = status  + "GTDNumber字段缺失!";
                }
                if ("".equals(gtdQuantity)){
                    status = status + "GTDQuantity字段缺失!";
                }
                if ("".equals(gtdAmount)){
                    status = status + "GTDAmount字段缺失!";
                }
                russiaGTD.setStatus(status);
            }
            excelOutput(excel_filePath,russiaGTD);
        }else {
            String msg = "OCR模板不适用";
            templateNot_GTD(companyCode, fileName, excel_filePath, msg);
        }
    }

    /**
     * 生成Excel表
     * @param excel_filePath
     * @param russiaGTD
     */
    public static void excelOutput(String excel_filePath, RussiaGTD russiaGTD) {
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        XSSFSheet xssfSheet = xssfWorkbook.createSheet("Sheet1");
        Row row0 = xssfSheet.createRow(0);
        String[] headers = new String[]{"FileName","CompanyCode", "GTDNumber", "GTDQuantity", "GTDAmount","OCRStatus"};
        for (int i = 0; i < headers.length; i++) {
            xssfSheet.setColumnWidth(i, 5000);
            XSSFCell cell = (XSSFCell) row0.createCell(i);
            XSSFRichTextString text = new XSSFRichTextString(headers[i]);
            cell.setCellValue(text);
        }
        int rowNum = 1;
        XSSFRow row = xssfSheet.createRow(rowNum);
        row.createCell(0).setCellValue(russiaGTD.getFileName());
        row.createCell(1).setCellValue(russiaGTD.getCompanyCode());
        row.createCell(2).setCellValue(russiaGTD.getGtdNumber());
        row.createCell(3).setCellValue(russiaGTD.getGtdQuantity());
        row.createCell(4).setCellValue(russiaGTD.getGtdAmount());
        row.createCell(5).setCellValue(russiaGTD.getStatus());
        try {
            FileOutputStream fileOutputStream = new FileOutputStream(excel_filePath);
            try {
                xssfWorkbook.write(fileOutputStream);
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }

    /**
     * 处理gtdQuantity数据
     * @param gtdQuantity
     * @return
     */
    public static String checkQuantity(String gtdQuantity){
        if (!"".equals(gtdQuantity)){
            Pattern p = Pattern.compile("[^.0-9]");//提取有效数字
            gtdQuantity = p.matcher(gtdQuantity).replaceAll("").trim();
        }
        return gtdQuantity;
    }

    /**
     * 模板不适用
     * @param companyCode
     * @param fileName
     * @param excel_filePath
     */
    public static void templateNot_GTD(String companyCode, String fileName,String excel_filePath, String message){
        RussiaGTD russiaGTD = new RussiaGTD();
        russiaGTD.setCompanyCode(companyCode);
        russiaGTD.setFileName(fileName);
        russiaGTD.setStatus(message);
        excelOutput(excel_filePath,russiaGTD);
    }
}
