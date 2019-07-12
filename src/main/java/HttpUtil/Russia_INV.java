package HttpUtil;

import DataClean.BaseUtil;
import DataClean.DataDO;
import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.sun.pdfview.PDFFile;
import com.sun.pdfview.PDFPage;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.ByteBuffer;
import java.nio.channels.FileChannel;
import java.util.*;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 该操作类实现首先将pdf转换为图片
 * 其次将转换的图片进行OCR扫描
 * 处理整合扫描获取到的数据
 *
 */
public class Russia_INV extends BaseUtil {
    public static void main(String[] args) {
//        String companyData = args[0];
//        String OCR_filePath = args[1];
//        String SAP_filePath = args[2];
//        String file_pdf = args[3];
        String companyData = "62F0-HQ-15112018-0.31";
        String OCR_filePath = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\俄罗斯海关单和发票\\4\\OCR_file.xlsx";//ocr所在路径
        String SAP_filePath = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\俄罗斯海关单和发票\\4\\SAP_file.xlsx";//sap 所在路径
        String file_pdf = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\俄罗斯海关单和发票\\4\\3798792 3798809 3798817 prealerts.pdf";//pdf 所在路径
        String file_img = file_pdf.substring(0,file_pdf.lastIndexOf("\\")) + "\\PDF_img";//转换成功的图片存储文件夹
        System.out.println("开始："+new Date());
        //pdf转换图片，将图片的路径集合返回
        List<String> imgUrl = (List<String>) splitPic(file_img, file_pdf);//返回图片路径集合
        String fileName = file_pdf.substring(file_pdf.lastIndexOf("\\") + 1);//截取图片路径
        try {
            clearINVFunction(companyData,OCR_filePath,SAP_filePath,file_img,fileName,imgUrl);
            System.out.println("结束："+new Date());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 将pdf转换成图片，保存文件夹
     *
     * @param file_img
     * @param file_pdf
     * @return
     */
    public static List<String> splitPic(String file_img, String file_pdf) {
        List<String> imgUrl = new ArrayList<>();
        try {
            String img_path = file_pdf.substring(file_pdf.lastIndexOf("\\") + 1, file_pdf.indexOf("."));
            File file = new File(file_pdf);
            RandomAccessFile raf = new RandomAccessFile(file, "r");
            FileChannel channel = raf.getChannel();
            ByteBuffer buf = channel.map(FileChannel.MapMode.READ_ONLY, 0, channel.size());
            PDFFile pdffile = new PDFFile(buf);

            String getPdfFilePath = file_img;
            //目录不存在，则创建目录
            File p = new File(getPdfFilePath);
            if (!p.exists()) {
                p.mkdir();
            }
            for (int i = 1; i <= pdffile.getNumPages(); i++) {
                PDFPage page = pdffile.getPage(i);
                int n = 6;
                Rectangle rect = new Rectangle(0, 0, (int) page.getBBox().getWidth(), (int) page.getBBox().getHeight());
                Image img = page.getImage(rect.width * n, rect.height * n, rect, null, true, true);

                BufferedImage tag = new BufferedImage(rect.width * n, rect.height * n, BufferedImage.TYPE_INT_RGB);
                tag.getGraphics().drawImage(img, 0, 0, rect.width * n, rect.height * n, null);

                //转换成功的图片路径dstName
                String dstName = getPdfFilePath + File.separator + img_path + "_" + i + ".jpg";
                FileOutputStream out = new FileOutputStream(dstName); // 输出到文件流
                String formatName = dstName.substring(dstName.lastIndexOf(".") + 1);
                ImageIO.write(tag, /*"GIF"*/ formatName /* format desired */, new File(dstName) /* target */);
                imgUrl.add(dstName);
                out.close();
            }
            System.out.println("pdf转换图片成功！");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return imgUrl;
    }

    /**
     *图片扫描，数据处理
     * @param companyData
     * @param OCR_filePath
     * @param SAP_filePath
     * @param file_img
     * @throws Exception
     */
    public static void clearINVFunction(String companyData,String OCR_filePath,String SAP_filePath,String file_img,String fileName, List<String> imgUrl) throws Exception {
        int index = companyData.indexOf("-");//获取第一次出现“-”的位置
        String companyCode = companyData.substring(0,companyData.indexOf("-",index + 1));
        String date =companyData.substring(companyData.indexOf("-",index + 1) + 1,companyData.indexOf("-",index + 1) + 9);
        String postingDate = date.substring(2,4) + "/" + date.substring(0,2) + "/" + date.substring(4);
        String exchangeRate = "";
        int count = 0;
        Pattern p = Pattern.compile("-");
        Matcher m = p.matcher(companyData);
        while (m.find()) {
            count++;
        }
        if (count == 3){
            exchangeRate = companyData.substring(companyData.lastIndexOf("-")+1);
        }
        if (imgUrl.size() > 0) {
            List<DataDO> dataDOList = new ArrayList<>();
            for (int i = 0; i <imgUrl.size() ;i ++){
                String invoiceReferenceNumber = "";
                String purchaseOrderNumber = "";
                String invoiceDate = "";
                String poShotText = "";
                String amountAndCurrency = "";
                String quantity = "";
                String amount = "";
                String currency = "";
                String file_path = imgUrl.get(i);//图片路径
                File file = new File(file_path);
                byte[] fileData = FileUtils.readFileToByteArray(file);//读取图片
                String json = ocrImageFile("APSigRusInv", file.getName(), fileData);
                System.out.println("输出OCR扫描json:"+ json);
                JSONObject jsonObject = JSON.parseObject(json);
                if (jsonObject.get("result").toString().contains("success")) {
                    //处理json数据
                    JSONArray jsonArray = jsonObject.getJSONObject("ocrResult").getJSONArray("ranges");
                    for (int j = 0; j <jsonArray.size(); j++) {
                        JSONObject object = (JSONObject) jsonArray.get(j);
                        switch (object.get("rangeId").toString()){
                            case "InvoiceReferenceNumber":
                                if (object.get("value")!= null){
                                    invoiceReferenceNumber = object.get("value").toString();
                                }else {
                                    invoiceReferenceNumber = "";
                                }
                                break;
                            case "PurchaseOrderNumber":
                                if (object.get("value") != null){
                                    purchaseOrderNumber = object.get("value").toString();
                                }else {
                                    purchaseOrderNumber = "";
                                }
                                break;
                            case "InvoiceDate":
                                if (object.get("value") != null){
                                    invoiceDate = object.get("value").toString();
                                }else {
                                    invoiceDate = "";
                                }
                                break;
                            case "POShotText":
                                if (object.get("value") != null){
                                    poShotText = object.get("value").toString();
                                }else {
                                    poShotText = "";
                                }
                                break;
                            case "AmountAndCurrency":
                                if (object.get("value") != null){
                                    amountAndCurrency =clearAmount(object.get("value").toString());
                                    String[] vl = amountAndCurrency.split(" ");
                                    List<String> list = Arrays.asList(vl);
                                    for (String str :list){
                                        System.out.println("shuhcu :" + str);
                                        if (str.matches("[a-zA-Z]{1,}")){//判断字符串全部是英文
                                            currency = str;
                                        }else {
                                            //判断字符串包含数字字符
                                            Pattern pa = Pattern.compile(".*\\d+.*");
                                            Matcher ma = pa.matcher(str);
                                            if (ma.matches()) {
                                                Pattern pp = Pattern.compile("[^0-9]");//提取有效数字
                                                amount = pp.matcher(str).replaceAll("").trim();
                                                amount = amount.substring(0, amount.length() - 2) + "." + amount.substring(amount.length() - 2);
                                            }
                                        }
                                    }
                                }
                                break;
                            case "Quantity":
                                if (object.get("value") != null){
                                    String[] vl = object.get("value").toString().split(" ");
                                    List<String> list = Arrays.asList(vl);
                                    for (String str :list){
                                        if (str.matches("[0-9]{1,}")){
                                            quantity = str;
                                        }
                                    }
                                }
                                break;
                        }
                    }
                    DataDO dataDO = new DataDO();
                    dataDO.setCompanyCode(companyCode);
                    dataDO.setFilepath(fileName);
                    dataDO.setDownloadStatus("OK");
                    dataDO.setAmount(convertData(amount));
                    dataDO.setQuantity(quantity);
                    dataDO.setPoshorttext(poShotText);
                    dataDO.setPurchaseOrderNumber(purchaseOrderNumber);
                    dataDO.setPostingDate(postingDate);
                    dataDO.setCurrency(currency);
                    dataDO.setInvoiceReferenceNumber(invoiceReferenceNumber);
                    dataDO.setInvoicedate(invoiceDate);
                    dataDO.setTaxCode("P0 (0% Input Tax for Goods,Services)");
                    dataDO.setExchangeRate(exchangeRate);
                    dataDO.setBaselineDate(invoiceDate);
                    dataDO.setAssignment(purchaseOrderNumber);
                    dataDO.setHeaderText(purchaseOrderNumber);
                    dataDO.setOCRStatus("PASS");
                    dataDOList.add(dataDO);
                }
            }
            //将不同的PO号拼接到一起
            Set set = new HashSet();
            String poNumber = "";
            for (DataDO dataDO : dataDOList){
                if (!set.contains(dataDO.getPurchaseOrderNumber())){
                    set.add(dataDO.getPurchaseOrderNumber());
                    poNumber = poNumber + dataDO.getPurchaseOrderNumber() + " ";
                }
            }
            System.out.println("输出poNumber:" + poNumber);
            String[] vl = poNumber.split(" ");
            List<String> list = Arrays.asList(vl);
            String poNumbers = "";
            for (int j = 0;j < list.size(); j++){
                String  num = list.get(j);
                if ( j >0 ){
                    num = num.substring(4);
                }
                poNumbers = poNumbers + num + "/";
            }
            poNumbers = poNumbers.substring(0,poNumbers.length() - 1);
            System.out.println("输出poNumbers:" + poNumbers);
            List<DataDO> dataDOS_OCR = new ArrayList<>();//存储所有获取到的数据
            List<DataDO> dataDOS_ASP = new ArrayList<>();//存储所有有效的数据
            List<DataDO> dataDOS_Exc = new ArrayList<>();//存储所有无效的数据（字段缺失、无法识别）
            for (DataDO dataDO : dataDOList){
                dataDO.setPurchaseOrderNumber(poNumber);
                dataDO.setAssignment(poNumbers);
                dataDO.setHeaderText(poNumbers);
                if ("".equals(dataDO.getInvoiceReferenceNumber()) || "".equals(dataDO.getInvoicedate()) || "".equals(dataDO.getPoshorttext())
                || "".equals(dataDO.getPurchaseOrderNumber()) || "".equals(dataDO.getAmount()) || "".equals(dataDO.getQuantity()) || "".equals(dataDO.getCurrency())){
                    dataDO.setOCRStatus("字段缺失");
                    dataDOS_Exc.add(dataDO);
                }
                dataDOS_OCR.add(dataDO);
            }
            if (dataDOS_Exc.size() == 0){
                for (DataDO dataDO : dataDOS_OCR){
                    dataDO.setOCRStatus("OK");
                }
                dataDOS_ASP.addAll(dataDOS_OCR);
                //生成SAP表
                if (dataDOS_ASP.size() > 0){
                    excelOutput_SAP(dataDOS_ASP, SAP_filePath);
                }
            }
            //生成OCR表
            excelOutput_OCR(dataDOS_OCR, OCR_filePath);
        }
        File dir = new File(file_img);
        File[] files = dir.listFiles();
        for (int i =0 ;i < files.length;i++){
            files[i].delete();//删除文件
        }
        dir.delete();//删除文件夹
    }

    /**
     * 生成SAP表格
     * @param dataDOS
     * @param filename
     */
    public static void excelOutput_SAP(List<DataDO> dataDOS, String filename){
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        XSSFSheet xssfSheet = xssfWorkbook.createSheet("Sheet1");
        Row row0 = xssfSheet.createRow(0);
        String[] headers = new String[]{"FileName", "CompanyCode", "Amount", "InvoiceDate", "InvoiceReferenceNumber", "InvoiceReferenceNumber2", "POShortText", "PurchaseOrderNumber", "Quantity", "TaxAmount", "TotalAmount", "UnitPrice", "Currency", "SONumber", "GoodsDescription", "OCRStatus", "PostingDate", "TaxCode", "Text", "BaselineDate", "ExchangeRate", "PaymentBlock", "Assignment", "HeaderText"};
        for (int i = 0; i < headers.length; i++) {
            XSSFCell cell = (XSSFCell) row0.createCell(i);
            xssfSheet.setColumnWidth(i, 5000);
            XSSFRichTextString text = new XSSFRichTextString(headers[i]);
            cell.setCellValue(text);
        }
        int rowNum = 1;
        for (DataDO dataDO : dataDOS) {
            XSSFRow row1 = xssfSheet.createRow(rowNum);
            row1.createCell(0).setCellValue(dataDO.getFilepath());
            row1.createCell(1).setCellValue(dataDO.getCompanyCode());
            row1.createCell(2).setCellValue(dataDO.getAmount());
            row1.createCell(3).setCellValue(dataDO.getInvoicedate());
            row1.createCell(4).setCellValue(dataDO.getInvoiceReferenceNumber());
            row1.createCell(5).setCellValue(dataDO.getInvoiceReferenceNumber2());
            row1.createCell(6).setCellValue(dataDO.getPoshorttext());
            row1.createCell(7).setCellValue(dataDO.getPurchaseOrderNumber());
            row1.createCell(8).setCellValue(dataDO.getQuantity());
            row1.createCell(9).setCellValue(dataDO.getTaxAmount());
            row1.createCell(10).setCellValue(dataDO.getTotalAmount());
            row1.createCell(11).setCellValue(dataDO.getUnitPrice());
            row1.createCell(12).setCellValue(dataDO.getCurrency());
            row1.createCell(13).setCellValue(dataDO.getSOnumber());
            row1.createCell(14).setCellValue(dataDO.getGoodDescription());
            row1.createCell(15).setCellValue(dataDO.getOCRStatus());
            row1.createCell(16).setCellValue(dataDO.getPostingDate());
            row1.createCell(17).setCellValue(dataDO.getTaxCode());
            row1.createCell(18).setCellValue(dataDO.getText());
            row1.createCell(19).setCellValue(dataDO.getInvoicedate());
            row1.createCell(20).setCellValue(dataDO.getExchangeRate());
            row1.createCell(21).setCellValue(dataDO.getPaymentBlock());
            row1.createCell(22).setCellValue(dataDO.getAssignment());
            row1.createCell(23).setCellValue(dataDO.getHeaderText());
            rowNum++;
        }
        try {
            FileOutputStream fileOutputStream = new FileOutputStream(filename);
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
     * Amount数据处理
     * @param data
     * @return
     */
    public static String clearAmount(String data){
        if (!"".equals(data)){
            data = data .replace("LSD","USD");
            data = data.replace("I ","U");
            data = data.replace("(SD","USD");
        }
        return data;
    }
}