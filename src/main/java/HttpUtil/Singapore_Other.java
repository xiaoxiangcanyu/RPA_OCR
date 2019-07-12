package HttpUtil;

import DataClean.BaseUtil;
import DataClean.DataDO;
import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 韩国以及泰国电产发票数据处理
 */
public class Singapore_Other extends BaseUtil {
    public static final String QUANTITY = "^[0-9]+[a-zA-Z]+([\\s\\S]*)+$";
    public static void main(String[] args)throws Exception{
        String companyCode = args[0];
        String filePath = args[1];
        String OCR_filePath = args[2];
        String SAP_filePath = args[3];
//        String companyCode="6560-HQ-PO12345556 PO3092";
//        String filePath = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AP\\新加坡其他国家\\泰国电产采购发票\\201905131923\\Thai-HQ\\COAU7043063750C0193DTH035WMPO3956(BEST)-0.jpg";
//        String OCR_filePath ="C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AP\\新加坡其他国家\\泰国电产采购发票\\201905131923\\Thai-HQ\\test\\OCR_file.xls";
//        String SAP_filePath= "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AP\\新加坡其他国家\\泰国电产采购发票\\201905131923\\Thai-HQ\\test\\SAP_file.xls";

        //泰国电产 PO号需从companyCode中获取，获取如下：
        String poNumber = getPONumber(companyCode);
        companyCode = companyCode.substring(0,companyCode.lastIndexOf("-"));
        System.out.println("输出companyCode ：" + companyCode);
        System.out.println("输出poNumber ：" + poNumber);
        String fileName = filePath.substring(filePath.lastIndexOf("\\") + 1);//截取图片路径
        //不同的companyCode需要不同的OCR模板
        try {
            switch (companyCode){
                case "6430-HQ"://韩国发票数据处理
                    clearROKFunction(companyCode, filePath, OCR_filePath, SAP_filePath, fileName);
                    break;
                case "6560-HQ"://泰国电产数据处理
                    clearTLFunction(companyCode, filePath, OCR_filePath, SAP_filePath, fileName,poNumber);
                    break;
            }
            System.out.println("运行结束！！！");
        }catch (Exception e){
            String msg = "OCR data incompleted";
            templateNot(companyCode, fileName, OCR_filePath, msg);
        }
    }
    /**
     * 韩国发票数据处理
     * @param companyCode
     * @param filePath
     * @param OCR_filePath
     * @param SAP_filePath
     * @param fileName
     * @throws Exception
     */
    public static void clearROKFunction(String companyCode, String filePath, String OCR_filePath, String SAP_filePath, String fileName) throws Exception{
        String invoiceReferenceNumber = "";
        String invoiceDate = "";
        String purchaseOrderNumber = "";
        String goodsDescription = "";
        String amountAndTotal = "";
        String totalAmount = "";
        String currency = "";
       File file = new File(filePath);
       byte[] fileData = FileUtils.readFileToByteArray(file);//读取图片
       //OCR模板扫描 返回结果数据
       String json = ocrImageFile("APSigKor", file.getName(), fileData);
       System.out.println("输出韩国json:"+ json);
       JSONObject jsonObject = JSON.parseObject(json);
       if (jsonObject.get("result").toString().contains("success")){//判断是否返回扫描成功数据
           //处理json数据
           JSONArray jsonArray = jsonObject.getJSONObject("ocrResult").getJSONArray("ranges");
           Map<String,Object> map = new HashMap<>();
           for (int i = 0; i <jsonArray.size(); i++){
               JSONObject object = (JSONObject) jsonArray.get(i);
               switch (object.get("rangeId").toString()){
                   case "InvoiceReferenceNumber":
                       if (object.get("value") != null) {
                           invoiceReferenceNumber =  object.get("value").toString();
                       }
                       break;
                   case "InvoiceDate":
                       if (object.get("value") != null){
                           invoiceDate = clearInvoiceDate(object.get("value").toString());
                       }
                       break;
                   case "PurchaseOrderNumber":
                       if (object.get("value") != null){
                           purchaseOrderNumber = object.get("value").toString();
                       }
                       break;
                   case "GoodsDescription":
                       if (object.get("value") != null){
                           goodsDescription = clearGoodsDescription(object.get("value").toString());
                       }
                       break;
                   case "POShortText":
                       if (object.get("value") != null){
                           String[] vl = object.get("value").toString().split(" ");
                           List<String> list = Arrays.asList(vl);
                           map.put("POShortText", list);
                       }else {
                           List<String> list = new ArrayList<>();
                           map.put("POShortText", list);
                       }
                       break;
                   case "Quantity":
                       if (object.get("value") != null){
                           String[] vl = object.get("value").toString().split(" ");
                           List<String> list = Arrays.asList(vl);
                           List<String> list1 = new ArrayList<>();
                           for (String quantity :list){
                               if (quantity.matches(QUANTITY)){
                                   Pattern pp = Pattern.compile("[^0-9]");
                                   quantity = pp.matcher(quantity).replaceAll("").trim();
                                   list1.add(quantity);
                               }
                           }
                           map.put("Quantity", list1);
                       }else {
                           List<String> list = new ArrayList<>();
                           map.put("Quantity", list);
                       }
                       break;
                       //模板扫描时，amount以及totalAmount同时获得，
                   case "Amount":
                       if (object.get("value") != null){
                           amountAndTotal = object.get("value").toString().replace(" ","");
                           System.out.println("输出amount:"+amountAndTotal);
                           String[] regerSplit = REGER_CURRENCY;//获取货币号数组
                           List<Double> lists = new ArrayList<>();
                           for (String curr : regerSplit){//遍历
                               if (amountAndTotal.contains(curr)){//判断amount 包含那个货币号
                                   currency = curr;//赋值
                                   String[] vl = amountAndTotal.split(curr);
                                   List<String> list = Arrays.asList(vl);//转成集合
                                   for (int j = 1; j < list.size(); j++){
                                       String amount = clearData(list.get(j));
                                       if (!"".equals(amount)){
                                           lists.add(Double.parseDouble(amount));
                                       }
                                   }
                               }
                           }
                           map.put("Amount",lists);
                       }else {
                           List<String> list = new ArrayList<>();
                           map.put("Amount", list);
                       }
                       break;
               }
           }
           System.out.println("输出currency："+currency);
           //获取amount集合，获取集合的最大值即为totalAmount，并将最大值删除
           List<Double> amountList = (List<Double>) map.get("Amount");
           if (amountList.size() > 0){
               System.out.println("输出最大值：" + Collections.max(amountList));
               totalAmount = String.valueOf(Collections.max(amountList));
               int j = amountList.lastIndexOf(Double.parseDouble(totalAmount));
               amountList.remove(j);
           }
           for (String key : map.keySet()) {//查看
               System.out.println(key + ":" + map.get(key));
           }
           System.out.println("输出invoiceDate：" + invoiceDate);
           List<String> POShortTextLit = (List<String>) map.get("POShortText");
           List<String> quantityList = (List<String>) map.get("Quantity");
           if (quantityList.size() > 0 && quantityList.size() == amountList.size() && quantityList.size() == POShortTextLit.size()){
               List<DataDO> dataDOList = new ArrayList<>();
               for (int i = 0 ; i < quantityList.size(); i++){
                   DataDO dataDO = new DataDO();
                   dataDO.setCompanyCode(companyCode);
                   dataDO.setFilepath(fileName);
                   dataDO.setDownloadStatus("OK");
                   dataDO.setInvoiceReferenceNumber(clearIRNumber(invoiceReferenceNumber));
                   dataDO.setInvoicedate(invoiceDate);
                   dataDO.setPurchaseOrderNumber(purchaseOrderNumber);
                   dataDO.setGoodDescription(goodsDescription);
                   dataDO.setTotalAmount(convertData(totalAmount));
                   dataDO.setAmount(convertData(String.valueOf(amountList.get(i))));
                   dataDO.setQuantity(quantityList.get(i));
                   //price的模板扫描异常，不采用，单独计算
                   if (!"".equals(amountList.get(i)) && !"".equals(quantityList.get(i))){
                       double d  =  amountList.get(i) / Double.parseDouble(quantityList.get(i));
                       dataDO.setUnitPrice(String.valueOf((double) Math.round(d * 100) / 100));
                   }else {
                       dataDO.setUnitPrice("");
                   }
                   dataDO.setPoshorttext(POShortTextLit.get(i));
                   dataDO.setTaxCode("X0 (no tax)");
                   dataDO.setPostingDate(clearPostingDate(invoiceDate));
                   dataDO.setCurrency(currency);
                   dataDO.setOCRStatus("PASS");
                   dataDOList.add(dataDO);
               }
               List<DataDO> dataDOS_OCR = new ArrayList<>();//存储所有获取到的数据
               List<DataDO> dataDOS_ASP = new ArrayList<>();//存储所有有效的数据
               List<DataDO> dataDOS_Exc = new ArrayList<>();//存储所有无效的数据（字段缺失、无法识别）
               for (DataDO dataDO : dataDOList) {
                   String status = "";//定义变量来保存字段无效信息
                   if (!"".equals(dataDO.getInvoiceReferenceNumber()) && !"".equals(dataDO.getInvoicedate()) && !"".equals(dataDO.getGoodDescription()) && !"".equals(dataDO.getPurchaseOrderNumber()) && !"".equals(dataDO.getTotalAmount()) && !"".equals(dataDO.getOCRStatus())){
                       if (dataDO.getInvoicedate().contains( " ")){
                           status = status + "InvoiceDate Cannot recognize";
                       }
                       //判断数据是否有字段缺失
                       if (!"".equals(status)){
                           dataDO.setOCRStatus(status);
                           dataDOS_Exc.add(dataDO);//将含有无效字段的数据存储起来
                       }
                       dataDOS_OCR.add(dataDO);
                   }else {
                       dataDO.setOCRStatus("Cannot detect");
                       dataDOS_Exc.add(dataDO);
                       dataDOS_OCR.add(dataDO);
                   }
               }
               //将有效数据保存到SAP集合
               if (dataDOS_Exc.size() > 0){
                   for (DataDO dataDO : dataDOS_Exc){
                       for (DataDO dataDO1 : dataDOS_OCR){
                           if (!dataDO1.getInvoiceReferenceNumber().equals(dataDO.getInvoiceReferenceNumber()) && !"".equals(dataDO1.getInvoiceReferenceNumber())){
                               dataDOS_ASP.add(dataDO1);
                           }
                       }
                   }
               }else {
                   dataDOS_ASP.addAll(dataDOS_OCR);
               }
               //判断是否可以生成ASP表
               System.out.println("输出sap数据数量" + dataDOS_ASP.size());
               if(dataDOS_ASP.size() > 0){
                   //遍历ASP表数据以及OCR表数据，将OCR表数据中可以生成ASP表的数据的OCRStatus修改为“OK”
                   for (DataDO dataDO : dataDOS_ASP){
                       for (DataDO dataDO1 : dataDOS_OCR){
                           if (dataDO1.getInvoiceReferenceNumber().equals(dataDO.getInvoiceReferenceNumber())){
                               dataDO1.setOCRStatus("OK");
                           }
                       }
                   }
                   //生成SAP表
                   excelOutput_SAP(dataDOS_ASP, SAP_filePath);
               }
               //生成OCR表
               excelOutput_OCR(dataDOS_OCR, OCR_filePath);
           }else {
               String msg = "OCR data incompleted";
               templateNot(companyCode, fileName, OCR_filePath, msg);
           }
       }else {
           String msg = "OCR Template not match";
           templateNot(companyCode, fileName, OCR_filePath, msg);
       }
    }

    /**
     * 过滤price
     * @param price
     * @return
     */
    public static String clearPrice(String price){
        if (!"".equals(price)){
            StringBuilder sb = new StringBuilder(price);
            String str = sb.substring(0,2);
            if ("IJ".equals(str)){
                sb.replace(0,2,"U");
            }
            price = sb.toString();
        }
        return price;
    }

    /**
     * 过滤goodsDescription
     * @param goodsDescription
     * @return
     */
    public static String clearGoodsDescription(String goodsDescription){
        if (!"".equals(goodsDescription)){
            StringBuilder sb = new StringBuilder(goodsDescription);
            String str = sb.substring(0,7);
            if ("GPRIGHT".equals(str)){
                sb.replace(0, 7, "UPRIGHT");
            }
            goodsDescription = sb.toString();
        }
        return goodsDescription;
    }

    /**
     * 处理postingDate
     * 若获取到的invoiceDate与校验发票的年月相同（当前年月），postingDate既为invoiceDate
     * 否者postingDate既为当前年月的当月第一天
     * @return
     */
    public static String clearPostingDate(String invoiceDate){
        String postingDate = "";
        if (!"".equals(invoiceDate) && !invoiceDate.contains(" ")){
            DateFormat dFormat = new SimpleDateFormat("MM/dd/yyyy");
            String startDate = invoiceDate.substring(0,2) + invoiceDate.substring(5);//获取invoiceDate年月
            System.out.println("发票年月："+ startDate);
            //获取当前年月
            Calendar calendar = Calendar.getInstance();
            calendar.setTime(new Date());
            String month = String.valueOf(calendar.get(Calendar.MONTH) +1);
            String year = String.valueOf(calendar.get(Calendar.YEAR));
            String date = month + "/" + year;
            System.out.println("当前年月："+ date);
            calendar.set(Calendar.DAY_OF_MONTH,1);
            //当前日期既为本月第一天
            String nowDate = dFormat.format(calendar.getTime());
            if (date.equals(startDate)){
                postingDate = invoiceDate;
            }else {
                postingDate = nowDate;
            }
        }
        return postingDate;
    }

    /**
     * 生成ASP表格
     * @param dataDOS
     * @param filename
     */
    public static void excelOutput_SAP(List<DataDO> dataDOS, String filename) throws Exception{
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        XSSFSheet xssfSheet = xssfWorkbook.createSheet("Sheet1");
        Row row0 = xssfSheet.createRow(0);
        String[] headers = new String[]{"FileName", "CompanyCode", "Amount", "InvoiceDate", "InvoiceReferenceNumber", "InvoiceReferenceNumber2", "POShortText", "PurchaseOrderNumber", "Quantity", "TaxAmount", "TotalAmount", "UnitPrice", "Currency", "SONumber", "GoodsDescription", "OCRStatus", "PostingDate", "TaxCode", "Text", "BaselineDate", "ExchangeRate", "PaymentBlock", "Assignment", "HeaderText"};
        for (int i = 0; i < headers.length; i++) {
            XSSFCell cell = (XSSFCell) row0.createCell(i);
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
            row1.createCell(15).setCellValue("OK");
            row1.createCell(16).setCellValue(dataDO.getPostingDate());
            row1.createCell(17).setCellValue(dataDO.getTaxCode());
            if ("6430-HQ".equals(dataDO.getCompanyCode())){
                row1.createCell(18).setCellValue(dataDO.getPurchaseOrderNumber() + "/" +dataDO.getGoodDescription());
            }else {
                row1.createCell(18).setCellValue(dataDO.getPurchaseOrderNumber());
            }
            row1.createCell(19).setCellValue(dataDO.getInvoicedate());
            row1.createCell(20).setCellValue("");
            row1.createCell(21).setCellValue("");
            row1.createCell(22).setCellValue("");
            row1.createCell(23).setCellValue("");
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
     * 泰国电产数据处理
     * @param companyCode
     * @param filePath
     * @param OCR_filePath
     * @param SAP_filePath
     * @param fileName
     * @throws Exception
     */
    public static void clearTLFunction(String companyCode, String filePath, String OCR_filePath, String SAP_filePath, String fileName, String poNumber) throws Exception{
        String invoiceReferenceNumber = "";
        String invoiceDate = "";
        String totalAmount = "";
        String currency = "";
        File file = new File(filePath);
        byte[] fileData = FileUtils.readFileToByteArray(file);//读取图片
        //OCR模板扫描 返回结果数据
        String json = ocrImageFile("APSigTha", file.getName(), fileData);
        System.out.println("输出泰国json:"+ json);
        JSONObject jsonObject = JSON.parseObject(json);
        if (jsonObject.get("result").toString().contains("success")) {//判断是否返回扫描成功数据
            //处理json数据
            JSONArray jsonArray = jsonObject.getJSONObject("ocrResult").getJSONArray("ranges");
            Map<String,Object> map = new HashMap<>();
            for (int i = 0; i <jsonArray.size(); i++){
                JSONObject object = (JSONObject) jsonArray.get(i);
                switch (object.get("rangeId").toString()){
                    case "InvoiceReferenceNumber":
                        if (object.get("value") != null) {
                            invoiceReferenceNumber =  object.get("value").toString();
                        }
                        break;
                    case "InvoiceDate":
                        if (object.get("value") != null){
                            invoiceDate = clearInvoiceDate(object.get("value").toString());
                        }
                        break;
                    case "POShortText":
                        if (object.get("value") != null){
                            String[] vl = object.get("value").toString().split(" ");
                            List<String> list = Arrays.asList(vl);
                            map.put("POShortText", list);
                        }else {
                            List<String> list = new ArrayList<>();
                            map.put("POShortText", list);
                        }
                        break;
                    case "Quantity":
                        if (object.get("value") != null){
                            String[] vl = object.get("value").toString().split(" ");
                            List<String> list = Arrays.asList(vl);
                            List<String> list1 = new ArrayList<>();
                            for (String quantity :list){
                                Pattern pattern = Pattern.compile("([0-9]{1,}PCS)|([0-9]{1,}SETS)");
                                Matcher matcher = pattern.matcher(quantity);
                                quantity = "";
                                while (matcher.find()){
                                    quantity = matcher.group();
                                }
                                System.out.println("第一步清洗之后quantity:"+quantity);
                                Pattern p = Pattern.compile(REGER);//提取有效数字
                                quantity = p.matcher(quantity).replaceAll("").trim();
                                if (!"".equals(quantity)){
                                    list1.add(quantity);
                                }
                            }
                            map.put("Quantity", list1);
                        }else {
                            List<String> list = new ArrayList<>();
                            map.put("Quantity", list);
                        }
                        break;
                    case "Price":
                        if (object.get("value") != null){
                            String unitPrice = object.get("value").toString().replace(" ","");
                            System.out.println("输出price：" + unitPrice);
                            unitPrice = unitPrice.replace("O","0");
                            String[] regerSplit = REGER_CURRENCY;//获取货币号数组
                            List<String> list1 = new ArrayList<>();
                            for (String curr : regerSplit){
                                if (unitPrice.contains(curr)){
                                    currency = curr;
                                    String[] vl = unitPrice.split(curr);
                                    List<String> list = Arrays.asList(vl);
                                    for (int j = 1; j < list.size(); j++){
                                        String price = clearData(list.get(j));
                                        if (!"".equals(price)){
                                            list1.add(price);
                                        }
                                    }
                                    break;
                                }
                            }
                            map.put("UnitPrice", list1);
                        }else {
                            List<String> list = new ArrayList<>();
                            map.put("UnitPrice", list);
                        }
                        break;
                    case "Amount":
                        if (object.get("value") != null){
                            String amountAndTotal = object.get("value").toString().replace(" ","");
                            System.out.println("输出amountAndTotal：" + amountAndTotal);
                            String[] regerSplit = REGER_CURRENCY;//获取货币号数组
                            List<Double> list1 = new ArrayList<>();
                            for (String curr : regerSplit){
                                if (amountAndTotal.contains(curr)) {
                                    String[] vl = amountAndTotal.split(curr);
                                    List<String> list = Arrays.asList(vl);
                                    for (int j = 1; j < list.size(); j++){
                                        String amount = clearData(list.get(j));
                                        if (!"".equals(amount)){
                                            list1.add(Double.parseDouble(amount));
                                        }
                                    }
                                    break;
                                }
                            }
                            map.put("Amount", list1);
                        }else {
                            List<String> list = new ArrayList<>();
                            map.put("Amount", list);
                        }
                        break;
                }
            }
            List<Double> amountList = (List<Double>) map.get("Amount");
            if (amountList.size() > 0){
                totalAmount = String.valueOf(Collections.max(amountList));//获取集合中的最大值
                int i = amountList.lastIndexOf(Double.parseDouble(totalAmount));//获取最大值最后出现位置
                System.out.println("最大值totalAmount：" + totalAmount);
                amountList.remove(i);
            }
            for (String key : map.keySet()) {
                System.out.println(key + ":" + map.get(key));
            }
            List<String> priceList = (List<String>) map.get("UnitPrice");
            List<String> quantityList = (List<String>) map.get("Quantity");
            List<String> POShortTextLit = (List<String>) map.get("POShortText");
            if (quantityList.size() > 0 && quantityList.size() == priceList.size() && quantityList.size() == amountList.size() && quantityList.size() == POShortTextLit.size()){
                List<DataDO> dataDOList = new ArrayList<>();
                for (int i = 0 ; i < quantityList.size(); i++){
                    DataDO dataDO = new DataDO();
                    dataDO.setCompanyCode(companyCode);
                    dataDO.setFilepath(fileName);
                    dataDO.setDownloadStatus("OK");
                    dataDO.setInvoiceReferenceNumber(clearIRNumber(invoiceReferenceNumber));
                    dataDO.setInvoicedate(invoiceDate);
                    dataDO.setPurchaseOrderNumber(poNumber);
                    dataDO.setTotalAmount(convertData(totalAmount));
                    dataDO.setAmount(convertData(String.valueOf(amountList.get(i))));
                    dataDO.setQuantity(quantityList.get(i));
                    dataDO.setUnitPrice(priceList.get(i));
                    dataDO.setPoshorttext(POShortTextLit.get(i));
                    dataDO.setTaxCode("V0 (Input VAT Receivable 0%)");
                    dataDO.setPostingDate(clearPostingDate2());
                    dataDO.setCurrency(currency);
                    dataDO.setOCRStatus("PASS");
                    dataDOList.add(dataDO);
                }
                List<DataDO> dataDOS_OCR = new ArrayList<>();//存储所有获取到的数据
                List<DataDO> dataDOS_ASP = new ArrayList<>();//存储所有有效的数据
                List<DataDO> dataDOS_Exc = new ArrayList<>();//存储所有无效的数据（字段缺失、无法识别）
                for (DataDO dataDO : dataDOList) {
                    String status = "";//定义变量来保存字段无效信息
                    if (!"".equals(dataDO.getInvoiceReferenceNumber()) && !"".equals(dataDO.getInvoicedate()) &&  !"".equals(dataDO.getPurchaseOrderNumber()) && !"".equals(dataDO.getTotalAmount())){
                        if (dataDO.getInvoicedate().contains( " ")){
                            status = status + "InvoiceDate Cannot recognize!";
                        }
                        //判断数据是否有字段缺失
                        if (!"".equals(status)){
                            dataDO.setOCRStatus(status);
                            dataDOS_Exc.add(dataDO);//将含有无效字段的数据存储起来
                        }
                        dataDOS_OCR.add(dataDO);
                    }else {
                        dataDO.setOCRStatus("Cannot detect");
                        dataDOS_OCR.add(dataDO);
                        dataDOS_Exc.add(dataDO);
                    }
                }
                //将有效数据保存到ASP集合
                if (dataDOS_Exc.size() > 0){
                    for (DataDO dataDO : dataDOS_Exc){
                        for (DataDO dataDO1 : dataDOS_OCR){
                            if (!dataDO1.getInvoiceReferenceNumber().equals(dataDO.getInvoiceReferenceNumber()) && !"".equals(dataDO1.getInvoiceReferenceNumber())){
                                dataDOS_ASP.add(dataDO1);
                            }
                        }
                    }
                }else {
                    dataDOS_ASP.addAll(dataDOS_OCR);
                }
                //判断是否可以生成ASP表
                System.out.println("输出sap数据数量" + dataDOS_ASP.size());
                if(dataDOS_ASP.size() > 0){
                    //遍历ASP表数据以及OCR表数据，将OCR表数据中可以生成ASP表的数据的OCRStatus修改为“OK”
                    for (DataDO dataDO : dataDOS_ASP){
                        for (DataDO dataDO1 : dataDOS_OCR){
                            if (dataDO1.getInvoiceReferenceNumber().equals(dataDO.getInvoiceReferenceNumber())){
                                dataDO1.setOCRStatus("OK");
                            }
                        }
                    }
                    //生成SAP表
                    excelOutput_SAP(dataDOS_ASP, SAP_filePath);
                }
                //生成OCR表
                excelOutput_OCR(dataDOS_OCR, OCR_filePath);
            }else {
                String msg = "OCR data incompleted";
                templateNot(companyCode, fileName, OCR_filePath, msg);
            }
        }else {
            String msg = "OCR template not matched";
            templateNot(companyCode, fileName, OCR_filePath, msg);
        }
    }
    /**
     * 获取postingDate
     * @return
     */
    public static String clearPostingDate2(){
        DateFormat dFormat = new SimpleDateFormat("MM/dd/yyyy");
        String date = dFormat.format(new Date());
        return date;
    }

    /**
     * 处理invoiceDate
     * @param invoiceDate
     * @return
     */
    public static String clearInvoiceDate(String invoiceDate){
        if (!"".equals(invoiceDate)){
            Pattern pp = Pattern.compile("[^0-9]");
            invoiceDate = pp.matcher(invoiceDate).replaceAll("").trim();
            if (invoiceDate.length() == 8) {
                invoiceDate = invoiceDate.substring(4,6) + "/" + invoiceDate.substring(6, 8) + "/" + invoiceDate.substring(0, 4);
            }else {
                invoiceDate = invoiceDate + " ";
            }
        }
        return invoiceDate;
    }

    /**
     * 处理PO号，标准的PO号为 56********
     * @param companyCode
     * @return
     */
    public static String getPONumber(String companyCode){
//        //获取companyCode中“-”的数量
//        int count = 0;
//        Pattern p = Pattern.compile("-");
//        Matcher m = p.matcher(companyCode);
//        while (m.find()) {
//            count++;
//        }
//        String poNumber = "";
//        if (count > 1){
//            poNumber = companyCode.substring(companyCode.lastIndexOf("-") + 1);
//        }
//        String number = "";
//        if (!"".equals(poNumber)){
//            //将获取到的PO号进行拆分
//            String[] vl = poNumber.split("PO");
//            List<String> list = Arrays.asList(vl);
//            for (String str : list){
//                System.out.println("切割之后的PO号："+str);
//                if (!"".equals(str)){
//                    //判断po号的位数，位数不够的需要在前面补“0”
//                    int index = str.length();
//                    StringBuffer sb = null;
//                    while (index < 8) {
//                        sb = new StringBuffer();
//                        sb.append("0").append(str);// 左补0
//                        str = sb.toString();
//                        index = str.length();
//                    }
//                    //最后在生成的PO号开头补存“56”
//                    number = number + " " + "56" + str;
//                }
//            }
//        }
//        return number.trim();
        String PONumberOriginal ="";
        String POSONumber= "";
        String POSONumberFinal = "";
        PONumberOriginal = companyCode.substring(8);
        companyCode = companyCode.trim().substring(0,7);
        System.out.println(companyCode);
        String[] vl =  PONumberOriginal.split("PO");
        List<String> list =Arrays.asList(vl);
        for(String str:list){
            if (!"".equals(str)) {
                System.out.println(str);
                POSONumber = "56" + str.trim();
                int ZeroLength = 10 - POSONumber.length();
                StringBuilder stringBuilder = new StringBuilder("");
                for (int i = 0; i < ZeroLength; i++) {
                    stringBuilder = stringBuilder.append("0");
                }
                String str1 = stringBuilder.toString();
                POSONumberFinal = "56" + str1 + str + " " + POSONumberFinal.trim();
            }
        }
        return POSONumberFinal;
    }
}
