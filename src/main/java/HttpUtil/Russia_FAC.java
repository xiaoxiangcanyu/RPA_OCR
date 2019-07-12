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
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Russia_FAC extends BaseUtil {
    public static void main(String[] args) throws Exception{
        String companyCode = args[0];
        String filePath1 = args[1];
        String filePath2 = args[2];
        String OCR_filePath = args[3];
        String SAP_filePath = args[4];
//        String companyCode="62F0-FAC-123456";
//        String filePath1 = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AP\\俄罗斯工厂\\62F0-FAC-fail\\90069578-0.jpg";
//        String filePath2 = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AP\\俄罗斯工厂\\62F0-FAC-fail\\90069578-1.jpg";
//        String OCR_filePath ="C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AP\\俄罗斯工厂\\62F0-FAC-fail\\OCR_file.xls";
//        String SAP_filePath= "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AP\\俄罗斯工厂\\62F0-FAC-fail\\SAP_file.xls";
        String fileName = filePath1.substring(filePath1.lastIndexOf("\\") + 1);//截取图片路径
//        try{
            clearRUSFunction(companyCode,filePath1,filePath2,OCR_filePath,SAP_filePath,fileName);
//        }catch (Exception e){
//            String error = "OCR data incompleted";
//            templateNot(companyCode, fileName, OCR_filePath, error);
//        }
    }

    /**
     * 数据处理
     * @param companyCode
     * @param filePath1
     * @param filePath2
     * @param OCR_filePath
     * @param SAP_filePath
     * @param fileName
     * @throws Exception
     */
    public static void clearRUSFunction(String companyCode, String filePath1,String filePath2, String OCR_filePath, String SAP_filePath, String fileName)throws Exception{
        File file1 = new File(filePath1);
        File file2 = new File(filePath2);
        byte[] fileData1 = FileUtils.readFileToByteArray(file1);
        byte[] fileData2 = FileUtils.readFileToByteArray(file2);
        String PurchaseOrderNumber = companyCode.substring(companyCode.lastIndexOf("-") + 1);
        String invoiceReferenceNumberAndDate = "";
        String invoiceReferenceNumberAndDate2 = "";
        String taxAmount = "";
        String taxAmount2 = "";
        String totalAmount = "";
        String totalAmount2 = "";
        String invoiceReferenceNumber = "";
        String invoiceReferenceNumber2 = "";
        String invoiceDate = "";
        String invoiceDate2 = "";
        String poShortText = "";
        String poShortText2 = "";
        String amount = "";
        String amount2 = "";
        String unitPrice = "";
        String unitPrice2 = "";
        String quantity = "";
        String quantity2 = "";
        double sumAmount = 0.0;
        double sumTaxAmount = 0.0;
        double totAmount = 0.0;
        double sumAmount2 = 0.0;
        double sumTaxAmount2 = 0.0;
        double totAmount2 = 0.0;
        //获取第一页的json数据
        String json1 = ocrImageFile("APRus01_0109-0624", file1.getName(), fileData1);
        JSONObject jsonObject1 = JSON.parseObject(json1);
        System.out.println("输出第一页的json数据："+json1);
        //获取第二页的json数据
        String json2 = ocrImageFile("APRus02_1016_0624", file2.getName(), fileData2);
        JSONObject jsonObject2 = JSON.parseObject(json2);
        System.out.println("输出第二页的json数据："+json2);

        //判断json模板是否返回成功
        if (jsonObject1.get("result").toString().contains("success") && jsonObject2.get("result").toString().contains("success")) {
            //处理json1数据
            JSONArray jsonArray1 = jsonObject1.getJSONObject("ocrResult").getJSONArray("ranges");
            System.out.println("+++++++"+jsonArray1);
            for (int i = 0; i <jsonArray1.size(); i++){
                JSONObject object1 = (JSONObject) jsonArray1.get(i);
                switch (object1.get("rangeId").toString()){
                    //InvoiceReferenceNumber And InvoiceDate整体获得，需要将单独的两个字段拆分出来
                    case "InvoiceReferenceNumberAndDate"://InvoiceReferenceNumber 和 invoiceReferenceNumber 一同获取
                        if (object1.get("value") != null){
                            invoiceReferenceNumberAndDate =object1.get("value").toString();
                        }
                        break;
                    case "TotalAmount":
                        if (object1.get("value") != null){
                            totalAmount = cleanTotalAmount(object1.get("value").toString()) ;
                        }
                        break;
                    case "TaxAmount":
                        if (object1.get("value") != null){
                            taxAmount = object1.get("value").toString() ;
                        }
                        break;
                    case "InvoiceReferenceNumber":
                        if (object1.get("value") != null){
                            invoiceReferenceNumber = checkInvNumber(object1.get("value").toString()) ;
                        }
                        break;
                    case "InvoiceDate":
                        if (object1.get("value") != null){
                            invoiceDate = checkDate(object1.get("value").toString());
                        }
                        break;
                }
            }
            //获取invoiceDate And invoiceReferenceNumber
            if (invoiceDate.contains(" ") || "".equals(invoiceDate)){
                Map<String,Object> map =handleNumberAndDate(invoiceReferenceNumberAndDate,"900");
                invoiceDate = map.get("invoiceDate").toString();
            }
            if (invoiceReferenceNumber.length() != 8 || "".equals(invoiceReferenceNumber)){
                Map<String,Object> map =handleNumberAndDate(invoiceReferenceNumberAndDate,"900");
                invoiceReferenceNumber = map.get("invoiceReferenceNumber").toString();
            }
            //处理json2数据
            JSONArray jsonArray2 = jsonObject2.getJSONObject("ocrResult").getJSONArray("ranges");
            for (int k = 0; k <jsonArray2.size(); k++){
                JSONObject object2 = (JSONObject) jsonArray2.get(k);
                switch (object2.get("rangeId").toString()){
                    case "TotalAmount":
                        if (object2.get("value") != null){
                            totalAmount2 =cleanTotalAmount(object2.get("value").toString());
                        }
                        break;
                    case "TaxAmount":
                        if (object2.get("value") != null){
                            taxAmount2 = cleanTaxAmount(object2.get("value").toString());
                        }
                        break;
                    case "InvoiceReferenceNumberAndDate":
                        if (object2.get("value") != null){
                            invoiceReferenceNumberAndDate2 = object2.get("value").toString() ;
                        }
                        break;
                    case "InvoiceReferenceNumber":
                        if (object2.get("value") != null){
                            invoiceReferenceNumber2 = checkInvNumber(object2.get("value").toString());
                        }
                        break;
                    case "InvoiceDate":
                        if (object2.get("value") != null){
                            invoiceDate2 = checkDate(object2.get("value").toString());
                        }
                        break;
                }
            }
            //获取invoiceDate2 And invoiceReferenceNumber2
            if (invoiceDate2.contains(" ") || "".equals(invoiceDate2)){
                Map<String,Object> map2 =handleNumberAndDate(invoiceReferenceNumberAndDate2,"800");
                invoiceDate2 = map2.get("invoiceDate").toString();
            }
            if (invoiceReferenceNumber2.length() != 8 || "".equals(invoiceReferenceNumber2)){
                Map<String,Object> map2 =handleNumberAndDate(invoiceReferenceNumberAndDate2,"800");
                invoiceReferenceNumber2 = map2.get("invoiceReferenceNumber").toString();
            }
            System.out.println("输出invoiceDate: "+invoiceDate);
            System.out.println("输出invoiceReferenceNumber: "+invoiceReferenceNumber);
            System.out.println("输出invoiceDate2: "+invoiceDate2);
            System.out.println("输出invoiceReferenceNumber2: "+invoiceReferenceNumber2);
            //获取表格数据
            JSONArray jsonArrayRowDatas1 = jsonArray1.getJSONObject(4).getJSONArray("rowDatas");
//            System.out.println("++++++++++++"+jsonArray2.getJSONObject(1));
            JSONArray jsonArrayRowDatas2 = jsonArray2.getJSONObject(1).getJSONArray("rowDatas");
            System.out.println("输出一图数据条数："+ jsonArrayRowDatas1.size());
            System.out.println("输出二图数据条数："+ jsonArrayRowDatas2.size());
            //判断两张发票获取到的数据条目数是否相等
            if (jsonArrayRowDatas1.size() == jsonArrayRowDatas2.size()){
                //整理第一张发票数据（若第一张发票有字段获取不到数据，就获取第二张发票对应的数据）
                List<DataDO> dataDOList = new ArrayList<>();
                for (int i = 0; i <jsonArrayRowDatas1.size(); i++){
                    JSONArray jsonArray = jsonArrayRowDatas1.getJSONArray(i);
                    JSONArray jsonArr2 = jsonArrayRowDatas2.getJSONArray(i);
                    for (int j = 0; j <jsonArray.size(); j++){
                        JSONObject object = (JSONObject) jsonArray.get(j);
                        JSONObject object2 = (JSONObject) jsonArr2.get(j);
                        switch (object.get("columnId").toString()){
                            case "Amount":
                                if (object.get("value") != null){
                                    amount = cleanAmount(object.get("value").toString());
                                }else {
                                    if (object2.get("value") != null){
                                        amount = cleanAmount(object2.get("value").toString()) ;
                                    }else {
                                        amount = "";
                                    }
                                }
                                break;
                            case "POShortText":
//                                System.out.println("断点处的POShortText："+object.get("value")+object.get("value") != null);
                                if (object.get("value") != null){
                                    poShortText = cleanPOShortText(object.get("value").toString());

                                }else {
                                    if (object2.get("value") != null){
                                        poShortText =cleanPOShortText(object2.get("value").toString());
                                    }else {
                                        poShortText = "";
                                    }
                                }
                                break;
                            case "UnitPrice":
                                if (object.get("value") != null){
                                    unitPrice = cleanPrice(object.get("value").toString());
                                }else {
                                    if (object2.get("value") != null){
                                        unitPrice = cleanPrice(object2.get("value").toString());
                                    }else {
                                        unitPrice = "";
                                    }
                                }
                                break;
                            case "Quantity"://若Quantity在两张发票都获取不到数据，通过计算的方式获取
                                if (object.get("value") != null){
                                    quantity = cleanQuantity(object.get("value").toString());
                                }else {
                                    if (object2.get("value") != null){
                                        quantity = cleanQuantity(object2.get("value").toString()) ;
                                    }else {
                                        if (!"".equals(amount) && !"".equals(unitPrice)){
                                            Double qutity = Double.parseDouble(amount) / Double.parseDouble(unitPrice);
                                            quantity =String.valueOf((double) Math.round(qutity * 100) / 100);
                                        }else {
                                            quantity = "";
                                        }
                                    }
                                }
                                break;
                        }
                    }
                    DataDO dataDO_original = new DataDO();
                    dataDO_original.setFilepath(fileName);
                    dataDO_original.setAmount(amount);
                    dataDO_original.setDownloadStatus("OK");
                    dataDO_original.setInvoiceReferenceNumber(invoiceReferenceNumber);
                    dataDO_original.setInvoiceReferenceNumber2(invoiceReferenceNumber2);
                    dataDO_original.setCompanyCode(companyCode);
                    dataDO_original.setInvoicedate(invoiceDate);
                    dataDO_original.setQuantity(quantity);
                    dataDO_original.setPoshorttext(poShortText);
                    dataDO_original.setUnitPrice(unitPrice);
                    dataDO_original.setPurchaseOrderNumber(PurchaseOrderNumber);
                    dataDO_original.setTaxAmount(cleanTaxAmount(taxAmount));
                    dataDO_original.setTotalAmount(cleanTotalAmount(totalAmount));
                    dataDO_original.setTaxCode("J2 (20% Input Tax for Goods (Deferred Tax))");
                    dataDO_original.setPostingDate(handlePostingDate(invoiceDate));
                    dataDO_original.setOCRStatus("PASS");
                    dataDOList.add(dataDO_original);
                }
                // 整理第二张发票的数据（若第二张发票有字段获取不到数据，就获取第一张发票对应字段的数据）
                List<DataDO> dataDOList2 = new ArrayList<>();
                for (int i = 0; i <jsonArrayRowDatas2.size(); i++){
                    JSONArray jsonArray = jsonArrayRowDatas1.getJSONArray(i);
                    JSONArray jsonArr2 = jsonArrayRowDatas2.getJSONArray(i);
                    for (int j = 0; j <jsonArr2.size(); j++){
                        JSONObject object = (JSONObject) jsonArray.get(j);
                        JSONObject object2 = (JSONObject) jsonArr2.get(j);
                        switch (object2.get("columnId").toString()){
                            case "POShortText":
                                if (object2.get("value") != null){
                                    poShortText2 = cleanPOShortText(object2.get("value").toString());
                                }else {
                                    if (object.get("value") != null){
                                        poShortText2 = cleanPOShortText(object.get("value").toString());
                                    }else {
                                        poShortText2 = "";
                                    }
                                }
                                break;
                            case "Amount":
                                if (object2.get("value") != null){
                                    amount2 = cleanAmount(object2.get("value").toString());
                                }else {
                                    if (object.get("value") != null){
                                        amount2 = cleanAmount(object.get("value").toString()) ;
                                    }else {
                                        amount2 = "";
                                    }
                                }
                                break;
                            case "UnitPrice":
                                if (object2.get("value") != null){
                                    unitPrice2 = cleanPrice(object2.get("value").toString());
                                }else {
                                    if (object.get("value") != null){
                                        unitPrice2 = cleanPrice(object.get("value").toString());
                                    }else {
                                        unitPrice2 = "";
                                    }
                                }
                                break;
                            case "Quantity"://若Quantity在两张发票都获取不到数据，通过计算的方式获取
                                if (object2.get("value") != null){
                                    quantity2 = cleanQuantity(object2.get("value").toString());
                                }else {
                                    if (object.get("value") != null){
                                        quantity2 = cleanQuantity(object.get("value").toString());
                                    }else {
                                        if (!"".equals(amount2) && !"".equals(unitPrice2)){
                                            Double qt = Double.parseDouble(amount2) / Double.parseDouble(unitPrice2);
                                            quantity2 = String.valueOf((double) Math.round(qt * 100) / 100);
                                        }else {
                                            quantity2 = "";
                                        }
                                    }
                                }
                                break;
                        }
                    }
                    DataDO dataDO_original = new DataDO();
                    dataDO_original.setFilepath(fileName);
                    dataDO_original.setAmount(amount2);
                    dataDO_original.setDownloadStatus("OK");
                    dataDO_original.setInvoiceReferenceNumber(invoiceReferenceNumber);
                    dataDO_original.setInvoiceReferenceNumber2(invoiceReferenceNumber2);
                    dataDO_original.setCompanyCode(companyCode);
                    dataDO_original.setInvoicedate(invoiceDate2);
                    dataDO_original.setQuantity(quantity2);
                    dataDO_original.setPoshorttext(poShortText2);
                    dataDO_original.setUnitPrice(unitPrice2);
                    dataDO_original.setPurchaseOrderNumber(PurchaseOrderNumber);
                    dataDO_original.setTaxAmount(cleanTaxAmount(taxAmount2));
                    dataDO_original.setTotalAmount(cleanTotalAmount(totalAmount2));
                    dataDO_original.setTaxCode("J2 (20% Input Tax for Goods (Deferred Tax))");
                    dataDO_original.setPostingDate(handlePostingDate(invoiceDate2));
                    dataDO_original.setOCRStatus("PASS");
                    dataDOList2.add(dataDO_original);
                }
                //第一张发票数据清洗，将存好的数据封装实体类
                boolean result = false;
                for (DataDO dataDO : dataDOList) {
                    //若获取到的amount为有效值，则计算所有amount的和
                    if (!dataDO.getAmount().contains(" ") && !"".equals(dataDO.getAmount())){
                        sumAmount = sumAmount + Double.parseDouble(dataDO.getAmount());
                    }else {
                        result = true;
                    }
                }
                //totalAmount 和 taxAmount在扫描过程中存在大量的识别不完整情况
                // 所以发票的totalAmount 和 taxAmount需要计算的到
                //taxAmount的值为所有amount之和*0.2
                sumTaxAmount = sumAmount * 0.2;//计算taxAmount
                //结果保留两位小数
                BigDecimal bd = new BigDecimal(sumTaxAmount);
                sumTaxAmount = bd.setScale(2, BigDecimal.ROUND_HALF_UP).doubleValue();

                //totalAmount的值为所有amount的在加上taxAmount的值
                totAmount = sumAmount + sumTaxAmount;//计算totalAmount
                BigDecimal bd2 = new BigDecimal(totAmount);
                totalAmount =  bd2.setScale(2, BigDecimal.ROUND_HALF_UP).toPlainString();

                //第二张发票数据清洗，将存好的数据封装实体类
                boolean result2 = false;
                for (DataDO dataDO2 : dataDOList2) {
                    if (!dataDO2.getAmount().contains(" ")){
                        sumAmount2 = sumAmount2 + Double.parseDouble(dataDO2.getAmount());
                    }else {
                        result2 = true;
                    }
                }
                //totalAmount 和 taxAmount在扫描过程中存在大量的识别不完整情况
                // 所以发票的totalAmount 和 taxAmount需要计算的到
                //taxAmount的值为所有amount之和*0.2
                sumTaxAmount2 = sumAmount2 * 0.2;//计算taxAmount
                //结果保留两位小数
                BigDecimal bd3 = new BigDecimal(sumTaxAmount2);
                sumTaxAmount2 = bd3.setScale(2, BigDecimal.ROUND_HALF_UP).doubleValue();
                //totalAmount的值为所有amount的在加上taxAmount的值
                totAmount2 = sumTaxAmount2 + sumAmount2;//计算totalAmount
                BigDecimal bd4= new BigDecimal(totAmount2);
                totalAmount2 = bd4.setScale(2, BigDecimal.ROUND_HALF_UP).toPlainString();

                System.out.println("输出totalAmount：" + totalAmount);
                System.out.println("输出totalAmount2：" + totalAmount2);
                //判断两张发票是否为同一单的发票。若两张发票得到的invoiceDate 和 totalAmount相等，既为两张发票为同一单的发票
                if (invoiceDate.equals(invoiceDate2) && totalAmount.equals(totalAmount2) && !"".equals(invoiceDate)){
                    List<DataDO> dataDOS_OCR = new ArrayList<>();//存储所有获取到的数据
                    List<DataDO> dataDOS_ASP = new ArrayList<>();//存储所有有效的数据
                    List<DataDO> dataDOS_Exc = new ArrayList<>();//存储所有无效的数据（字段缺失、无法识别）
                    for (DataDO dataDO : dataDOList) {
                        dataDO.setTaxAmount(String.valueOf(sumTaxAmount));
                        dataDO.setTotalAmount(totalAmount);
                        String status = "";//定义变量来保存字段缺失信息
                        if (!"".equals(dataDO.getInvoicedate()) && !"".equals(dataDO.getInvoiceReferenceNumber()) && !"".equals(dataDO.getInvoiceReferenceNumber2()) &&!"".equals(dataDO.getAmount()) &&  !"".equals(dataDO.getUnitPrice()) && !"".equals(dataDO.getQuantity()) && !"".equals(dataDO.getPoshorttext())){
                            if (dataDO.getUnitPrice().contains(" ")){
                                status = status + "unitPrice Cannot recognize!";
                            }
                            if (dataDO.getQuantity().contains(" ")){
                                status = status + "quantity Cannot recognize!";
                            }
                            if (dataDO.getInvoiceReferenceNumber().contains(" ")){
                                status = status + "InvoiceReferenceNumber Cannot recognize!";
                            }
                            if (dataDO.getInvoiceReferenceNumber2().contains(" ")){
                                status = status + "InvoiceReferenceNumber2 Cannot recognize!";
                            }
                            if (dataDO.getPoshorttext().contains(" ")){
                                status = status + "POShortText Cannot recognize!";
                            }
                            //判断数据是否有字段缺失
                            if (!"".equals(status)){
                                dataDO.setOCRStatus(status);
                                dataDOS_Exc.add(dataDO);//将含有无效字段的数据存储起来
                            }
                            dataDOS_OCR.add(dataDO);
                        }else {
                            dataDO.setOCRStatus("Cannot detect!");
                            dataDOS_Exc.add(dataDO);//用来存储字段无效识别的数据
                            dataDOS_OCR.add(dataDO);
                        }
                    }
                    //将有效数据保存到ASP集合
                    if (dataDOS_Exc.size() > 0){
                        for (DataDO dataDO : dataDOS_Exc){
                            for (DataDO dataDO1 : dataDOS_OCR){
                                if (!dataDO1.getInvoiceReferenceNumber().equals(dataDO.getInvoiceReferenceNumber())&& !"".equals(dataDO1.getInvoiceReferenceNumber())){
                                    dataDOS_ASP.add(dataDO1);
                                }
                            }
                        }
                    }else {
                        dataDOS_ASP.addAll(dataDOS_OCR);
                    }
                    //判断是否可以生成SAP表
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
                    String str = "";
                    if (!invoiceDate.equals(invoiceDate2)){
                        str = str + "Invoice date not matched";
                    }
                    if (!totalAmount.equals(totalAmount2)){
                        if (result){
                            str = str + "Amount of Invoice 1 incompleted";
                        }
                        if (result2){
                            str = str + "Amount of Invoice 2 incompleted";
                        }
                        if (result == false && result2 == false){
                            str = str + "Total Amount not matched";
                        }
                    }
                    List<DataDO> dataDOLists = new ArrayList<>();
                    DataDO dataDO = new DataDO();
                    dataDO.setDownloadStatus("OK");
                    dataDO.setCompanyCode(companyCode);
                    dataDO.setFilepath(fileName);
                    dataDO.setOCRStatus(str);
                    dataDOLists.add(dataDO);
                    excelOutput_OCR(dataDOLists, OCR_filePath);
                }
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
     * 清洗invoiceReferenceNumber
     * @param invoiceReferenceNumber
     * @return
     */
    public static String cleanInvoiceReferenceNumber(String invoiceReferenceNumber){
        invoiceReferenceNumber = invoiceReferenceNumber.replace("j","1");
        invoiceReferenceNumber = invoiceReferenceNumber.replace("？","1");
        invoiceReferenceNumber = invoiceReferenceNumber.replace("?","1");
        return invoiceReferenceNumber;
    }
    /**
     * 清洗taxAmoun
     * @param taxAmount
     * @return
     */
    public static String cleanTaxAmount(String taxAmount){
        if (!"".equals(taxAmount)){
            if ("| ".equals(taxAmount.substring(0,2))){
                taxAmount = taxAmount.replace("| ", "");
            }
            if ("1 ".equals(taxAmount.substring(0,2))){
                taxAmount = taxAmount.replace("1 ", "");
            }
            taxAmount = taxAmount.replace("|", "");
            taxAmount = taxAmount.replace("\"", "");
            taxAmount = taxAmount.replace("”", "");
            taxAmount = taxAmount.replace("-", "");
            taxAmount = taxAmount.replace("“", "");
            taxAmount = taxAmount.replace("^", "");
            taxAmount = taxAmount.replace("S", "5");
            taxAmount = taxAmount.replace("l", "1");
            taxAmount = taxAmount.replaceAll("(?:I|!)", "1");
            taxAmount = taxAmount.replace("(", "1");
            taxAmount = taxAmount.replace(")", "1");
            taxAmount = taxAmount.replace("（", "1");
            taxAmount = taxAmount.replace("）", "1");
            taxAmount = taxAmount.replace("&", "8");
            taxAmount = taxAmount.replace("^", "");
            taxAmount = taxAmount.replace("g", "8");
            taxAmount = taxAmount.replace("o", "0");
            taxAmount = taxAmount.replace("O", "0");
            taxAmount = taxAmount.replace("Q", "0");
            taxAmount = taxAmount.replace("（", "1");
            taxAmount = taxAmount.replace("）", "1");
            taxAmount = taxAmount.replace("|", "1");
            taxAmount = taxAmount.replace("\\", "1");
            taxAmount = taxAmount.replace("/", "1");
            taxAmount = taxAmount.replace("\n", "");
            taxAmount = taxAmount.replace("↓", "1");
            taxAmount = taxAmount.replace("S", "8");
            taxAmount = taxAmount.replace(",", ".");
            taxAmount = taxAmount.replace("，", ".");
            taxAmount = taxAmount.replace(" ", "");
            try {
                taxAmount = taxAmount.substring(0,taxAmount.indexOf(".")) + taxAmount.substring(taxAmount.indexOf("."),taxAmount.indexOf(".") + 3);
                Float taxAmount_parse = Float.parseFloat(taxAmount);
            } catch (Exception e) {
                taxAmount = taxAmount + " ";
            }
        }else {
            taxAmount = " ";
        }
        return taxAmount;
    }

    /**
     * 清洗totalamount
     * @param totalamount
     * @return
     */
    public static String cleanTotalAmount(String totalamount){
        if (!"".equals(totalamount)){
            if ("| ".equals(totalamount.substring(0,2))){
                totalamount = totalamount.replace("| ", "");
            }
            if ("1 ".equals(totalamount.substring(0,2))){
                totalamount = totalamount.replace("1 ", "");
            }
            totalamount = totalamount.replace("|", "");
            totalamount = totalamount.replace("：", "");
            totalamount = totalamount.replace("^", "");
            totalamount = totalamount.replace(",", ".");
            totalamount = totalamount.replace("，", ".");
            totalamount = totalamount.replace("-", "");
            totalamount = totalamount.replace("S", "5");
            totalamount = totalamount.replace("I", "1");
            totalamount = totalamount.replace("l", "1");
            totalamount = totalamount.replaceAll("(?:I|!)", "1");
            totalamount = totalamount.replace("(", "1");
            totalamount = totalamount.replace(")", "1");
            totalamount = totalamount.replace("（", "1");
            totalamount = totalamount.replace("）", "1");
            totalamount = totalamount.replace("&", "8");
            totalamount = totalamount.replace("o", "0");
            totalamount = totalamount.replace("O", "0");
            totalamount = totalamount.replace("（", "1");
            totalamount = totalamount.replace("）", "1");
            totalamount = totalamount.replace("|", "1");
            totalamount = totalamount.replace("\\", "1");
            totalamount = totalamount.replace("/", "1");
            totalamount = totalamount.replace("\n", "");
            totalamount = totalamount.replace("↓", "1");
            totalamount = totalamount.replace(":","");
            totalamount = totalamount.replace(" ", "");
            try {
                totalamount = totalamount.substring(0,totalamount.indexOf(".")) + totalamount.substring(totalamount.indexOf("."),totalamount.indexOf(".") + 3);
                Float totalamount_parse =Float.parseFloat(totalamount);
            } catch (Exception e) {
                totalamount = totalamount + " ";
            }
        }else {
            totalamount = " ";
        }
        return totalamount;
    }

    /**
     * 清洗amount
     * @param amount
     * @return
     */
    public static String cleanAmount(String amount){
        System.out.println("输出amount结果："+ amount);
        if (!"".equals(amount)){
            if ("| ".equals(amount.substring(0,2))){
                amount = amount.replace("| ", "");
            }
            if ("1 ".equals(amount.substring(0,2))){
                amount = amount.replace("1 ", "");
            }

            amount =amount.toUpperCase();
            amount =amount.replace(",◦〔",".00");
            amount =amount.replace(".◦〔",".00");
            amount = amount.replace(",O〔",".00");
            amount = amount.replace(".O〔",".00");
            amount = amount.replace(",0〔",".00");
            amount =amount.replace(",O（",".00");
            amount = amount.replace(",0（",".00");
            amount =amount.replace(".O（",".00");
            amount = amount.replace(".0（",".00");
            amount = amount.replace(".0(",".00");
            amount = amount.replace(".0C",".00");
            amount = amount.replace(".O(",".00");
            amount = amount.replace(".OC",".00");
            amount = amount.replace(",0(",".00");
            amount = amount.replace(",0C",".00");
            amount = amount.replace(",O(",".00");
            amount = amount.replace(",OC",".00");
            amount = amount.replace(",", ".");
            amount = amount.replace("，", ".");
            amount = amount.replace(" ", ".");
            amount = amount.replace("S", "5");
            amount = amount.replace("I", "1");
            amount = amount.replace("l", "1");
            amount = amount.replaceAll("(?:I|!)", "1");
            amount = amount.replace("(", "1");
            amount = amount.replace(")", "1");
            amount = amount.replace("（", "1");
            amount = amount.replace("）", "1");
            amount = amount.replace("J", "1");
            amount = amount.replace("^", "");
            amount = amount.replace("&", "8");
            amount = amount.replace("g", "8");
            amount = amount.replace("o", "0");
            amount = amount.replace("O", "0");
            amount = amount.replace("Q", "0");
            amount = amount.replace("（", "1");
            amount = amount.replace("）", "1");
            amount = amount.replace("|", "1");
            amount = amount.replace("\\", "1");
            amount = amount.replace("/", "1");
            amount = amount.replace("\n", "");
            amount = amount.replace("↓", "1");
            amount = amount.replace("—", "");
            amount = amount.replace("一", "");
            amount = amount.replace("■", "");
            amount = amount.replace("-", "");
            amount = amount.replace(" ", "");
            try {
                amount = amount.substring(0,amount.indexOf(".")) + amount.substring(amount.indexOf("."),amount.indexOf(".") + 3);
                Float amount_parse = Float.parseFloat(amount);
            } catch (Exception e) {
                amount = amount + " ";
            }
        }
        return amount;
    }

    /**
     * 清洗price
     * @param price
     * @return
     */
    public static String cleanPrice(String price){
//        System.out.println("数据清理之前price:"+price);
        if (!"".equals(price)){
            if ("| ".equals(price.substring(0,2))){
                price = price.replace("| ", "");
            }
            if ("1 ".equals(price.substring(0,2))){
                price = price.replace("1 ", "");
            }
            price = price.replace(",", ".");
            price = price.replace("，", ".");
            price = price.replace("^", "");
            price = price.replace("S", "5");
            price = price.replace("I", "1");
            price = price.replace("l", "1");
            price = price.replaceAll("(?:I|!)", "1");
            price = price.replace("(", "1");
            price = price.replace(")", "1");
            price = price.replace("（", "1");
            price = price.replace("）", "1");
            price = price.replace("&", "8");
            price = price.replace("o", "0");
            price = price.replace("O", "0");
            price = price.replace("（", "1");
            price = price.replace("）", "1");
            price = price.replace("|", "1");
            price = price.replace("\\", "1");
            price = price.replace("/", "1");
            price = price.replace("\n", "");
            price = price.replace("↓", "1");
            price = price.replace("\n", "");
            price = price.replace(" ", "");
//            System.out.println("在拼接之前之前price:"+price);
            try {
                price = price.substring(0,price.indexOf(".")) + price.substring(price.indexOf("."),price.indexOf(".") + 3);
                System.out.println("输出结果price："+price);
                double price_parse = Double.parseDouble(price);
            } catch (Exception e) {
                price = price + " ";
            }
        }
        return price;
    }


        /**
         *清洗quantity
         * @param quantity
         * @return
         */
    public static String cleanQuantity(String quantity){
        if (!"".equals(quantity)){
            quantity = quantity.replace("-", "");
            quantity = quantity.replace("|", "");
            quantity = quantity.replace("丨", "");
            quantity = quantity.replace("—", "");
            quantity = quantity.replace("~", "");
            quantity = quantity.replace("，", "");
            quantity = quantity.replace(",", "");
            quantity = quantity.replace("_", "");
            quantity = quantity.replace("?", "");
            quantity = quantity.replace("？", "");
            quantity = quantity.replace("j", "");
            quantity = quantity.replace("J", "");
            quantity = quantity.replace("I", "");
            quantity = quantity.replace("i", "");
            quantity = quantity.replace(":", "");
            quantity = quantity.replace("：", "");
            quantity = quantity.replace("一", "");
            quantity = quantity.replace("O", "0");
            quantity = quantity.replace(" ", "");
            Pattern p = Pattern.compile("[^0-9]");
            quantity = p.matcher(quantity).replaceAll("").trim();
            try {
                Float quantity_parse = Float.parseFloat(quantity);
            } catch (Exception e) {
                quantity = quantity + " ";
            }
        }
        return  quantity;

    }
    /**
     *  清洗poShortText
     *  获取字符串类似于H*K的子串，获取子串位置之后的子串作为poShortText
     * @param poShortText
     * @return
     */
    public static String cleanPOShortText(String poShortText) {
        if (!"".equals(poShortText)){
            poShortText = poShortText.replace(" ","");
            poShortText = poShortText.replace("_","");
            poShortText = poShortText.replace("\\/","V");
            Pattern p = Pattern.compile("H[^H]*K");
            Matcher m = p.matcher(poShortText);
            List<String> stringList = new ArrayList<>();
            //获取符合表达式(H*K)的子串
            while(m.find()){
                stringList.add(m.group());
            }
            if (stringList.size() >0){
                poShortText = poShortText.substring(poShortText.lastIndexOf(stringList.get(stringList.size() - 1)) + stringList.get(stringList.size() - 1).length());
            }else if (stringList.size() < 0){
                poShortText="";
            }
        }else {
            poShortText = "";
        }
         return poShortText;
    }

    /**
     * 提取invoiceReferenceNumberAndDate And invoiceDate
     * 参数 index 是不同发票单号的前三位
     * @param invoiceReferenceNumberAndDate
     * @param index
     * @return
     */
    public static Map handleNumberAndDate(String invoiceReferenceNumberAndDate,String index){
        invoiceReferenceNumberAndDate = invoiceReferenceNumberAndDate.replace("J","1");
        Map<String,Object> map = new HashMap<>();
        String invoiceReferenceNumber = "";
        String invoiceDate = "";
        if (!"".equals(invoiceReferenceNumberAndDate)){
            //去除字符串中的空格
            invoiceReferenceNumberAndDate = invoiceReferenceNumberAndDate.replace(" ","");
            //提取字符串中的数字字符
            Pattern p = Pattern.compile("[^0-9]");
            String  str = p.matcher(invoiceReferenceNumberAndDate).replaceAll("").trim();//将字符创中的数字字符提取出来
            //找到“900或800”所在的位置，将invoiceReferenceNumber（单号）提取出来
            try {
                invoiceReferenceNumber = cleanInvoiceReferenceNumber(str.substring(str.indexOf(index),str.indexOf(index)+ 8));

            }catch (Exception e){
                invoiceReferenceNumber ="";
            }
            //将字符串中的单号去除
            str = str.replace(invoiceReferenceNumber,"");

            //获取当前年份
            Calendar calendar = Calendar.getInstance();
            calendar.setTime(new Date());
            String year_now = String.valueOf(calendar.get(Calendar.YEAR));
            //获取当前日期减去一年
            Calendar calendar1 = Calendar.getInstance();
            calendar1.setTime(new Date());
            calendar1.add(Calendar.YEAR, -1);//当前时间减去一年，即一年前的时间
            String year_old = String.valueOf(calendar1.get(Calendar.YEAR));
            //截取年份的前三位字符 如："2018"-"201"
            String year = year_now.substring(0,3);
            String date = "";
            String time = "";
            //使用截取到的年份字符在str中截取日期
            if (str.contains(year)){
                date = str.substring(0,str.lastIndexOf(year));
                time = year_now;
                //若当前年不能作为截取依据，就用当前年的前一年截取
                if (date.length() <= 0){
                    year = year_old.substring(0,3);
                    date = str.substring(0,str.lastIndexOf(year));
                    if (date.length() > 0){
                        time = year_old;
                    }
                }
            }
            //提取子串中的数字字符，去除符号
            try{
                //最后生成有效的发票单号
                date = date.substring(date.length()-4) + time;
                System.out.println("输出日期:" + date);
                System.out.println("date的长度:"+date.length());
                if (date.length() == 8) {
                    invoiceDate = date.substring(2,4) + "/" + date.substring(0, 2) + "/" + date.substring(4, 8);
                    //获取当前日期与发票日期进行比较，若发票日期大于当前日期，那么发票的年份为前一年
                    SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
                    Date nowDate = new Date();//获取当前系统时间
                    Date invDate = sdf.parse(invoiceDate);
                    if (invDate.getTime() > nowDate.getTime()){
                        invoiceDate = date.substring(2,4) + "/" + date.substring(0, 2) + "/" + time;
                    }
                }
            }
            catch (Exception e){
                invoiceDate = "";
                System.out.println("发票单号或时间错误！");
            }
        }
        map.put("invoiceDate",invoiceDate);
        map.put("invoiceReferenceNumber",invoiceReferenceNumber);
        return map;
    }

    /**
     * 处理
     * @param invDate
     * @return
     */
    public static String  checkDate(String invDate){
        if (!"".equals(invDate)){
            Pattern pp = Pattern.compile("[^0-9]");
            invDate = pp.matcher(invDate).replaceAll("").trim();
            if (invDate.length() == 8){
                invDate = invDate.substring(2,4) + "/" + invDate.substring(0, 2) + "/" + invDate.substring(4, 8);
            }else {
                invDate = invDate + " ";
            }
        }
        return invDate;
    }

    /**
     * 处理invNumber
     * @param invnumber
     * @return
     */
    public static String  checkInvNumber(String invnumber){
        if (!"".equals(invnumber)){
            Pattern pp = Pattern.compile("[^0-9]");
            invnumber = pp.matcher(invnumber).replaceAll("").trim();
        }
        return invnumber;
    }

        /**
         * 获取PostingDate
         * 若发票的校验日期与发票日期是同年月，PostingDate既为发票日期invoiceDate
         * 若非同一年月，PostingDate既为校验发票时间的本月第一天
         * @param invoiceDate
         * @return
         */
    public static String handlePostingDate(String invoiceDate)throws Exception{
        String postingDate = "";
        if (!"".equals(invoiceDate) && !invoiceDate.contains(" ")) {
            DateFormat dFormat = new SimpleDateFormat("MM/dd/yyyy");
            SimpleDateFormat sdf = new SimpleDateFormat("MM/yyyy");
            String startDate = invoiceDate.substring(0,2) + invoiceDate.substring(5);//获取invoiceDate年月
            System.out.println("发票年月："+ startDate);
            Date date1 = sdf.parse(startDate);
            //获取当前年月
            Calendar calendar = Calendar.getInstance();
            calendar.setTime(new Date());
            String month = String.valueOf(calendar.get(Calendar.MONTH) +1);
            String year = String.valueOf(calendar.get(Calendar.YEAR));
            String date = month + "/" + year;
            System.out.println("当前年月："+ date);
            Date date2 = sdf.parse(date);
            calendar.set(Calendar.DAY_OF_MONTH,1);//当前日期既为本月第一天
            String nowDate = dFormat.format(calendar.getTime());
            System.out.println("当前月第一天："+ dFormat.format(calendar.getTime()));
            if (date1.getTime() == date2.getTime()){
                System.out.println("#######同年月" );
                postingDate = invoiceDate;
            }else {
                postingDate = nowDate;
            }
        }
        return postingDate;
    }

    /**
     * 生成SAP表格
     * @param dataDOS
     * @param fileName
     */
    public static void excelOutput_SAP(List<DataDO> dataDOS, String fileName) throws Exception{
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        XSSFSheet xssfSheet = xssfWorkbook.createSheet("Sheet1");
        String[] headers = new String[]{"FileName", "CompanyCode", "Amount", "InvoiceDate", "InvoiceReferenceNumber", "InvoiceReferenceNumber2", "POShortText", "PurchaseOrderNumber", "Quantity", "TaxAmount", "TotalAmount", "UnitPrice", "Currency", "SONumber", "GoodsDescription", "OCRStatus", "PostingDate", "TaxCode", "Text", "BaselineDate", "ExchangeRate", "PaymentBlock", "Assignment", "HeaderText"};
        Row row0 = xssfSheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            XSSFCell cell = (XSSFCell) row0.createCell(i);
            xssfSheet.setColumnWidth(i, 4000);
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
            row1.createCell(12).setCellValue("RUB");
            row1.createCell(13).setCellValue(dataDO.getSOnumber());
            row1.createCell(14).setCellValue(dataDO.getGoodDescription());
            row1.createCell(15).setCellValue("OK");
            row1.createCell(16).setCellValue(dataDO.getPostingDate());
            row1.createCell(17).setCellValue(dataDO.getTaxCode());
            row1.createCell(18).setCellValue(dataDO.getPurchaseOrderNumber());
            row1.createCell(19).setCellValue(dataDO.getInvoicedate().replace("\n", ""));
            row1.createCell(20).setCellValue("");
            row1.createCell(21).setCellValue("");
            row1.createCell(22).setCellValue(dataDO.getInvoiceReferenceNumber2());
            row1.createCell(23).setCellValue(dataDO.getInvoiceReferenceNumber());
            rowNum++;
        }
        FileOutputStream fileOutputStream = new FileOutputStream(fileName);
        xssfWorkbook.write(fileOutputStream);

    }
}

