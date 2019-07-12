package HttpUtil;

import DataClean.BaseUtil;
import DataClean.DataDO;
import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.regex.Pattern;

public class Thailand_pdf extends BaseUtil{
    public static void main(String[] args) {
        String companyCode = args[0];
        String filePath = args[1];
        String OCR_filepath = args[2];
        String SAP_filepath = args[3];
        String filePath1 = args[4];
//        String companyCode = "6560-FAC";
//        String filePath = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AP\\新加坡用horis的国家\\泰国工厂\\泰国工厂AP\\test\\6560-FAC-1901276301.jpg";
//        String OCR_filepath = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AP\\新加坡用horis的国家\\泰国工厂\\泰国工厂AP\\test\\OCR.xls";
//        String SAP_filepath = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AP\\新加坡用horis的国家\\泰国工厂\\泰国工厂AP\\test\\SAP.xls";
//        String filePath1 = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AP\\新加坡用horis的国家\\泰国工厂\\泰国工厂AP\\GR WMS 2018-Dev 18（泰国Ap对照表）.xlsx";
        //从路径中获取companycode和filename
        String fileName = filePath.substring(filePath.lastIndexOf("\\") + 1);//截取图片路径
        try {
            clearTLFunction(companyCode, filePath, OCR_filepath, SAP_filepath, fileName, filePath1);
        } catch (Exception e) {
            String msg = "OCR data incompleted";
            templateNot(companyCode, fileName, OCR_filepath, msg);
        }

    }
    /**
     * 数据处理
     *
     * @param companyCode
     * @param filePath
     * @param OCR_filePath
     * @param SAP_filPath
     * @param fileName
     * @throws Exception
     */
    public static void clearTLFunction(String companyCode, String filePath, String OCR_filePath, String SAP_filPath, String fileName, String filePath1) throws Exception {
        File file = new File(filePath);
        byte[] fileData = FileUtils.readFileToByteArray(file);
        double sumAmount = 0.0;
        double taxAmount = 0.0;
        double totAmount = 0.0;
        int amountNum = 0;
        int quantityNum = 0;
        int priceNum = 0;
        int posNum = 0;
        String TaxAmount = "";
        String TotalAmount = "";
        String Amount = "";
        String UnitPrice = "";
        String InvoiceReferenceNumber = "";
        String InvoiceDate = "";
        String PurchaseOrderNumber = "";
        String POShortText = "";
        String Quantity = "";
        String TaxAndTotalAmount = "";
        //获取json
        String json = ocrImageFile("APTha_new_0318", file.getName(), fileData);
        System.out.println("json" + json);
        JSONObject jsonObject = JSON.parseObject(json);
        //判断json数据是否返回成功
        if (jsonObject.get("result").toString().contains("success")) {
            JSONArray jsonArray = jsonObject.getJSONObject("ocrResult").getJSONArray("ranges");
            //获取税单前几个字段
            for (int i = 0; i < jsonArray.size(); i++) {
                JSONObject object = (JSONObject) jsonArray.get(i);
                switch (object.get("rangeId").toString()) {
                    case "InvoiceReferenceNumber":
                        if (object.get("value") != null) {
                            InvoiceReferenceNumber = cleanInvoiceReferenceNumber(object.get("value").toString());
                        } else {
                            InvoiceReferenceNumber = "";
                        }
                        break;
                    case "InvoiceDate":
                        if (object.get("value") != null) {
                            InvoiceDate = cleanInvoiceDate(object.get("value").toString());
                        } else {
                            InvoiceDate = "";
                        }
                        break;
                    case "PurchaseOrderNumber":
                        if (object.get("value") != null) {
                            PurchaseOrderNumber = cleanPurchaseOrderNumber(object.get("value").toString());
                        } else {
                            PurchaseOrderNumber = "";
                        }
                        break;
                }
            }
            System.out.println("InvoiceReferenceNumber:"+InvoiceReferenceNumber);
            System.out.println("InvoiceDate:"+InvoiceDate);
            System.out.println();
            //判断发票号是否一致。ocr扫描发票号与图片路径发票号判断一致
                JSONArray jsonArrayRowDatas = jsonArray.getJSONObject(2).getJSONArray("rowDatas");
                System.out.println(jsonArrayRowDatas.size());
                //遍历table获取table里面的数据
                DateFormat dFormat = new SimpleDateFormat("MM/dd/yyyy");
                List<DataDO> dataDOList = new ArrayList<>();
                for (int i = 0; i < jsonArrayRowDatas.size(); i++) {
                    JSONArray jsonArray1 = jsonArrayRowDatas.getJSONArray(i);
                    System.out.println("循环第"+(i+1)+"行");
                    for (int j = 0; j < jsonArray1.size(); j++) {
//                        System.out.println("第"+(i+1)+"行共有"+jsonArray1.size()+"个元素");
                        JSONObject object = (JSONObject) jsonArray1.get(j);
                        switch (object.get("columnId").toString()) {

                            case "POShortText":
                                if (object.get("value") != null) {
                                    POShortText = object.get("value").toString();
                                } else {
                                    POShortText = "";
                                }
                                break;
                            case "UnitPrice":
                                if (object.get("value") != null) {
                                    UnitPrice = cleanPrice(object.get("value").toString());
                                } else {
                                    UnitPrice = "";
                                }
                                break;
                            case "Quantity":
                                if (object.get("value") != null) {
                                    Quantity = cleanQuantity(object.get("value").toString());
                                } else {
                                    Quantity = "";
                                }
                                break;
                            case "Amount":
                                if (object.get("value") != null) {
                                    Amount = cleanAmount(object.get("value").toString());
                                } else {
                                    Amount = "";
                                }
                                break;
                        }
                    }
                    System.out.println("第"+(i+1)+"行的数据:"+"POShortText:"+POShortText+",UnitPrice:"+UnitPrice+",Quantity:"+Quantity+",Amount:"+Amount);

                    DataDO dataDO_original = new DataDO();
                    dataDO_original.setFilepath(fileName);
                    dataDO_original.setAmount(Amount);
                    dataDO_original.setInvoiceReferenceNumber(InvoiceReferenceNumber);
                    dataDO_original.setCompanyCode(companyCode);
                    dataDO_original.setInvoicedate(InvoiceDate);
                    dataDO_original.setQuantity(Quantity);
                    dataDO_original.setPoshorttext(POShortText);
                    dataDO_original.setUnitPrice(UnitPrice);
                    dataDO_original.setPurchaseOrderNumber(PurchaseOrderNumber);
                    dataDO_original.setTaxAmount(TaxAmount);
                    dataDO_original.setTotalAmount(TotalAmount);
                    dataDO_original.setTaxCode("V7 (Input VAT Receivable 7%)");
                    dataDO_original.setPostingDate(dFormat.format(new Date()));
                    dataDOList.add(dataDO_original);
                }
                //数据清洗，将存好的数据封装实体类
                List<DataDO> dataDOList1 = new ArrayList<>();//用来存储清洗之后的数据
                for (DataDO dataDO : dataDOList) {
                    String Amount_clean = dataDO.getAmount();
                    if (!"".equals(Amount_clean) && !Amount_clean.equals(" ")) {//判断清洗后的Amount是否有效
                        sumAmount = sumAmount + Double.parseDouble(Amount_clean);
                        amountNum = amountNum + 1;//统计amount个数
                    }
                    String UnitPrice_clean = dataDO.getUnitPrice();
                    if (!"".equals(UnitPrice_clean)) {
                        priceNum = priceNum + 1;
                    }
                    if (!"".equals(dataDO.getQuantity())) {
                        quantityNum = quantityNum + 1;
                    }
                    String TaxAmount_clean = dataDO.getTaxAmount();
                    String TotalAmount_clean = dataDO.getTotalAmount();
                    DataDO dataDO1 = new DataDO();
                    dataDO1.setCompanyCode(companyCode);
                    dataDO1.setFilepath(dataDO.getFilepath());
                    dataDO1.setTotalAmount(TotalAmount_clean);
                    dataDO1.setAmount(Amount_clean);
                    dataDO1.setTaxAmount(TaxAmount_clean);
                    dataDO1.setPurchaseOrderNumber(dataDO.getPurchaseOrderNumber());
                    dataDO1.setInvoiceReferenceNumber(dataDO.getInvoiceReferenceNumber());
                    dataDO1.setUnitPrice(UnitPrice_clean);
                    String poShortText = cleanShortText(dataDO);
                    if (!"".equals(poShortText)) {
                        posNum = posNum + 1;//统计poShortText个数
                    }
                    dataDO1.setPoshorttext(poShortText);
                    dataDO1.setQuantity(dataDO.getQuantity());
                    dataDO1.setInvoicedate(dataDO.getInvoicedate());
                    dataDO1.setTaxCode(dataDO.getTaxCode());
                    dataDO1.setPostingDate(dataDO.getPostingDate());
                    dataDO1.setDownloadStatus("OK");
                    dataDO1.setOCRStatus("PASS");
                    dataDOList1.add(dataDO1);
                }
                System.out.println("输出：" + amountNum + "," + priceNum + "," + posNum + "," + quantityNum);

                //totalAmount 和 taxAmount在扫描过程中存在大量的识别不完整情况
                // 所以发票的totalAmount 和 taxAmount需要计算的到
                //taxAmount的值为所有amount之和*0.07
                taxAmount = sumAmount * 0.07;
                //结果保留两位小数
                BigDecimal bd2 = new BigDecimal(taxAmount);
                TaxAmount = bd2.setScale(2, BigDecimal.ROUND_HALF_UP).toPlainString();
                System.out.println("输出taxAmount的值：" + TaxAmount);
                //totalAmount的值为所有amount的在加上taxAmount的值
                totAmount = sumAmount + bd2.setScale(2, BigDecimal.ROUND_HALF_UP).doubleValue();
                BigDecimal bd3 = new BigDecimal(totAmount);
                TotalAmount = bd3.setScale(2, BigDecimal.ROUND_HALF_UP).toPlainString();
                System.out.println("输出TotalAmount的值：" + TotalAmount);
                //对于数据无法清洗的字段进行过滤和筛选
                List<DataDO> dataDOS_OCR = new ArrayList<>();//存储所有获取到的数据
                List<DataDO> dataDOS_ASP = new ArrayList<>();//存储所有有效的数据
                List<DataDO> dataDOS_Exc = new ArrayList<>();//存储所有无效的数据（字段缺失、无法识别）
                //判断获取到的所有字段的数据个数是否相等（OCR模板有没有扫描不到的数据）
                if (amountNum > 0 && amountNum == posNum && amountNum == priceNum && amountNum == quantityNum) {
                    for (DataDO dataDO : dataDOList1) {//遍历所有处理好的数据
                        dataDO.setTotalAmount(TotalAmount);
                        dataDO.setTaxAmount(TaxAmount);
                        String status = "";
                        if (!"".equals(dataDO.getInvoicedate()) && !"".equals(dataDO.getPurchaseOrderNumber()) && !"".equals(dataDO.getInvoiceReferenceNumber())) {
                            if (dataDO.getInvoiceReferenceNumber().contains(" ")) {
                                status = status + "InvoiceReferenceNumber Cannot Recognize!";
                            }
                            if (dataDO.getInvoicedate().contains(" ")) {
                                status = status + "InvoiceDate Cannot Recognize!";
                            }
                            if (dataDO.getQuantity().contains(" ")) {
                                status = status + "Quantity Cannot Recognize!";
                            }
                            if (dataDO.getUnitPrice().contains(" ")) {
                                status = status + "UnitPrice Cannot Recognize!";
                            }
                            if (dataDO.getPurchaseOrderNumber().contains(" ")) {
                                status = status + "PurchaseOrderNumber Cannot Recognize!";
                            }

                            if (!"".equals(status)) {
                                dataDO.setOCRStatus(status);
                                dataDOS_Exc.add(dataDO);//用来存储字段无效的数据
                            }
                            dataDOS_OCR.add(dataDO);
                        } else {
                            dataDO.setOCRStatus("Cannot Detect");
                            dataDOS_OCR.add(dataDO);
                            dataDOS_Exc.add(dataDO);//用来存储字段无效识别的数据
                        }
                    }
                    //将有效数据保存到ASP集合
                    if (dataDOS_Exc.size() > 0) {
                        for (DataDO dataDO : dataDOS_Exc) {
                            for (DataDO dataDO1 : dataDOS_OCR) {
                                if (!dataDO1.getInvoiceReferenceNumber().equals(dataDO.getInvoiceReferenceNumber()) && !"".equals(dataDO1.getInvoiceReferenceNumber())) {
                                    dataDOS_ASP.add(dataDO1);
                                }
                            }
                        }
                    } else {
                        dataDOS_ASP.addAll(dataDOS_OCR);
                    }
                    System.out.println("输出sap数据数量" + dataDOS_ASP.size());
                    //判断是否可生成ASP表
                    if (dataDOS_ASP.size() > 0) {
                        //遍历SAP表数据以及OCR表数据，将OCR表数据中可以生成SAP表的数据的OCRStatus修改为“OK”
                        for (DataDO dataDO : dataDOS_ASP) {
                            for (DataDO dataDO1 : dataDOS_OCR) {
                                if (dataDO1.getInvoiceReferenceNumber().equals(dataDO.getInvoiceReferenceNumber())) {
                                    dataDO1.setOCRStatus("OK");
                                }
                            }
                        }
                        //获取item参照表的表数据
                        List<DataDO> dataDoList = getFilePathDate(filePath1);
                        for (DataDO dataDO : dataDOS_ASP) {
                            String str = "";
                            for (DataDO dataDo1 : dataDoList){
                                if (dataDO.getInvoiceReferenceNumber().equals(dataDo1.getInvoicenum())){
                                    str = str + dataDo1.getItem() + " ";
                                }
                            }
                            dataDO.setItem(str);
                        }
                        //生成SAP表
                        excelOutput_SAP(dataDOS_ASP, SAP_filPath);
                    }
                    //生成OCR表
                    excelOutput_OCR(dataDOS_OCR, OCR_filePath);
                } else {
                    String msg = "OCR Data Incompleted";
                    templateNot(companyCode, fileName, OCR_filePath, msg);
                }

        } else {
            String msg = "OCR Data Incompleted";
            templateNot(companyCode, fileName, OCR_filePath, msg);
        }
    }

    /**
     * 清洗price业务
     *
     * @param price
     * @return
     */
    public static String cleanPrice(String price) {
//        System.out.println("unitPrice:" + price);
        if (!"".equals(price)) {
            price = price.replace(" ", "");
            price = price.replace(",", "");
            price = price.replace(".", "");
            price = price.replace("-", "");
            price = price.replace("，", "");
            price = price.replace(";", "");
            price = price.replace("；", "");
            price = price.replace(":", "");
            price = price.replace("S", "5");
            price = price.replace("I", "1");
            price = price.replace("?", "");
            price = price.replaceAll("(?:I|!)", "1");
            price = price.replace("(", "1");
            price = price.replace(")", "1");
            price = price.replace("（", "1");
            price = price.replace("）", "1");
            price = price.replace("J", "1");
            price = price.replace("j", "1");
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
            price = price.replace("^", "7");
            price = price.replace("•", "");
            price = price.replace("^", "7");
            price = price.substring(0, price.length() - 2) + "." + price.substring(price.length() - 2);
            price = price.replace("\n", "");
            try {
                Float price_parse = Float.parseFloat(price);
            } catch (Exception e) {
                price = price + " ";
            }
        } else {
            price = "";
        }
        return price;

    }


    /**
     * 清洗amount
     *
     * @param amount
     * @return
     */
    public static String cleanAmount(String amount) {
        if (!"".equals(amount)) {
            amount = amount.replace(" ", "");
            amount = amount.replace(",", "");
            amount = amount.replace(":", "");
            amount = amount.replace("-", "");
            amount = amount.replace(".", "");
            amount = amount.replace("，", "");
            amount = amount.replace("S", "5");
            amount = amount.replace("I", "1");
            amount = amount.replaceAll("(?:I|!)", "1");
            amount = amount.replace("(", "1");
            amount = amount.replace(")", "1");
            amount = amount.replace("（", "1");
            amount = amount.replace("）", "1");
            amount = amount.replace("J", "1");
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
            amount = amount.replace("^", "");
            amount = amount.substring(0, amount.length() - 2) + "." + amount.substring(amount.length() - 2);
            try {
                Float amount_parse = Float.parseFloat(amount);
            } catch (Exception e) {
                amount = amount + " ";
            }
        }
        return amount;

    }

    /**
     * 清洗invoiceDate业务
     *
     * @param invoiceDate
     * @return
     */
    public static String cleanInvoiceDate(String invoiceDate) {
        String invoiceDate_clean = "";
        if (!"".equals(invoiceDate)) {
            invoiceDate = invoiceDate.replace(",", "");
            invoiceDate = invoiceDate.replace(" ", "");
            invoiceDate = invoiceDate.replace(".", "");
            invoiceDate = invoiceDate.replace("S", "5");
            invoiceDate = invoiceDate.replace("I", "1");
            invoiceDate = invoiceDate.replace("J", "1");
            invoiceDate = invoiceDate.replace("l", "1");
            invoiceDate = invoiceDate.replaceAll("(?:I|!)", "1");
            invoiceDate = invoiceDate.replace("(", "1");
            invoiceDate = invoiceDate.replace(")", "1");
            invoiceDate = invoiceDate.replace("（", "1");
            invoiceDate = invoiceDate.replace("）", "1");
            invoiceDate = invoiceDate.replace("&", "8");
            invoiceDate = invoiceDate.replace("o", "0");
            invoiceDate = invoiceDate.replace("O", "0");
            invoiceDate = invoiceDate.replace("（", "1");
            invoiceDate = invoiceDate.replace("）", "1");
            invoiceDate = invoiceDate.replace("|", "1");
            invoiceDate = invoiceDate.replace("\\", "1");
            invoiceDate = invoiceDate.replace("/", "1");
            invoiceDate = invoiceDate.replace("\n", "");
            invoiceDate = invoiceDate.replace("↓", "1");
            if (invoiceDate.length() == 8) {
                String invoiceDate_month = invoiceDate.substring(2, 4);
                if (Integer.parseInt(invoiceDate_month) >= 1 && Integer.parseInt(invoiceDate_month) <= 11) {
                    Calendar calendar = Calendar.getInstance();
                    calendar.setTime(new Date());
                    String year = String.valueOf(calendar.get(Calendar.YEAR));
                    invoiceDate_clean = invoiceDate.substring(0, 4) + year;
                } else {
                    invoiceDate_clean = invoiceDate;
                }
                invoiceDate_clean = invoiceDate_clean.substring(2, 4) + "/" + invoiceDate_clean.substring(0, 2) + "/" + invoiceDate_clean.substring(4, 8);
                invoiceDate_clean = invoiceDate_clean.replace("\n", "");
                invoiceDate_clean = invoiceDate_clean.replace("\n", "");
                try {
                    SimpleDateFormat format = new SimpleDateFormat("MM/dd/yyyy");
                    format.setLenient(false);
                    format.parse(invoiceDate_clean);
                } catch (ParseException e) {
                    invoiceDate_clean = invoiceDate_clean + " ";
                }
            } else {
                invoiceDate_clean = invoiceDate_clean + " ";
            }
        }
        return invoiceDate_clean;
    }

    /**
     * 清洗invoiceReferenceNumber业务
     *
     * @param invoiceReferenceNumber
     * @return
     */
    public static String cleanInvoiceReferenceNumber(String invoiceReferenceNumber) {
        if (!"".equals(invoiceReferenceNumber)) {
            invoiceReferenceNumber = invoiceReferenceNumber.replace(" ", "");
            invoiceReferenceNumber = invoiceReferenceNumber.replace("-", "");
            invoiceReferenceNumber = invoiceReferenceNumber.replace(",", "");
            invoiceReferenceNumber = invoiceReferenceNumber.replace(".", "");
            invoiceReferenceNumber = invoiceReferenceNumber.replace("S", "5");
            invoiceReferenceNumber = invoiceReferenceNumber.replace("B", "13");
            invoiceReferenceNumber = invoiceReferenceNumber.replace("I", "1");
            invoiceReferenceNumber = invoiceReferenceNumber.replace("J", "1");
            invoiceReferenceNumber = invoiceReferenceNumber.replace("^", "7");
            invoiceReferenceNumber = invoiceReferenceNumber.replaceAll("(?:I|!)", "1");
            invoiceReferenceNumber = invoiceReferenceNumber.replace("(", "1");
            invoiceReferenceNumber = invoiceReferenceNumber.replace(")", "1");
            invoiceReferenceNumber = invoiceReferenceNumber.replace("（", "1");
            invoiceReferenceNumber = invoiceReferenceNumber.replace("）", "1");
            invoiceReferenceNumber = invoiceReferenceNumber.replace("&", "8");
            invoiceReferenceNumber = invoiceReferenceNumber.replace("o", "0");
            invoiceReferenceNumber = invoiceReferenceNumber.replace("O", "0");
            invoiceReferenceNumber = invoiceReferenceNumber.replace("Q", "0");
            invoiceReferenceNumber = invoiceReferenceNumber.replace("U", "0");
            invoiceReferenceNumber = invoiceReferenceNumber.replace("u", "0");
            invoiceReferenceNumber = invoiceReferenceNumber.replace("[", "1");
            invoiceReferenceNumber = invoiceReferenceNumber.replace("]", "1");
            invoiceReferenceNumber = invoiceReferenceNumber.replace("|", "1");
            invoiceReferenceNumber = invoiceReferenceNumber.replace("\\", "1");
            invoiceReferenceNumber = invoiceReferenceNumber.replace("/", "1");
            invoiceReferenceNumber = invoiceReferenceNumber.replace("\n", "");
            invoiceReferenceNumber = invoiceReferenceNumber.replace("↓", "1");
            invoiceReferenceNumber = invoiceReferenceNumber.trim();
            if (invoiceReferenceNumber.length() == 9 && invoiceReferenceNumber.substring(0, 1).equals("9")) {
                invoiceReferenceNumber = "1" + invoiceReferenceNumber;
            }
            try {
                Float INum = Float.parseFloat(invoiceReferenceNumber);
            } catch (Exception e) {
                invoiceReferenceNumber = invoiceReferenceNumber + " ";
            }
        }
        return invoiceReferenceNumber;
    }

    /**
     * 清洗purchaseOrderNumber业务
     *
     * @param purchaseOrderNumber
     * @return
     */
    public static String cleanPurchaseOrderNumber(String purchaseOrderNumber) {
        String purchaseOrderNumber_clean = "";
        if (!"".equals(purchaseOrderNumber)) {
            try {
                purchaseOrderNumber = purchaseOrderNumber.replace("(Fcb)", "");
                purchaseOrderNumber = purchaseOrderNumber.replace("(Feb)", "");
                purchaseOrderNumber = purchaseOrderNumber.replace("%", "56");
                purchaseOrderNumber = purchaseOrderNumber.replace("'", "");
                purchaseOrderNumber = purchaseOrderNumber.replace(",", "");
                purchaseOrderNumber = purchaseOrderNumber.replace(".", "");
                purchaseOrderNumber = purchaseOrderNumber.replace("，", "");
                purchaseOrderNumber = purchaseOrderNumber.replace("=", "");
                purchaseOrderNumber = purchaseOrderNumber.replace("■", "");
                purchaseOrderNumber = purchaseOrderNumber.replace("S", "5");
                purchaseOrderNumber = purchaseOrderNumber.replace("$", "5");
                purchaseOrderNumber = purchaseOrderNumber.replace("I", "1");
                purchaseOrderNumber = purchaseOrderNumber.replaceAll("(?:I|!)", "1");
                purchaseOrderNumber = purchaseOrderNumber.replace("（", "1");
                purchaseOrderNumber = purchaseOrderNumber.replace("）", "1");
                purchaseOrderNumber = purchaseOrderNumber.replace("^", "7");
                purchaseOrderNumber = purchaseOrderNumber.replace("&", "8");
                purchaseOrderNumber = purchaseOrderNumber.replace("o", "0");
                purchaseOrderNumber = purchaseOrderNumber.replace("O", "0");
                purchaseOrderNumber = purchaseOrderNumber.replace("(", "1");
                purchaseOrderNumber = purchaseOrderNumber.replace(")", "1");
                purchaseOrderNumber = purchaseOrderNumber.replace("|", "1");
                purchaseOrderNumber = purchaseOrderNumber.replace("\\", "1");
                purchaseOrderNumber = purchaseOrderNumber.replace("/", "1");
                purchaseOrderNumber = purchaseOrderNumber.replace("\n", "");
                purchaseOrderNumber = purchaseOrderNumber.replace("↓", "1");
                purchaseOrderNumber = purchaseOrderNumber.replace("l", "1");
                purchaseOrderNumber = purchaseOrderNumber.replace("8J", "81");
                purchaseOrderNumber = purchaseOrderNumber.replace(" ", "");
                System.out.println("po号 ："+ purchaseOrderNumber);
                boolean result2 = Pattern.matches("\\d{10}", purchaseOrderNumber.substring(0, 10));
                boolean result3 = purchaseOrderNumber.substring(0, 9).matches("\\d{9}");
                if (purchaseOrderNumber.contains("TH")) {
                    boolean result = purchaseOrderNumber.substring(6).substring(0, 10).matches("[0-9]+");
                    boolean result1 = purchaseOrderNumber.substring(6).substring(0, 9).matches("[0-9]+");
                    if (result == true && purchaseOrderNumber.substring(6).substring(0, 2).equals("56")) {
                        purchaseOrderNumber_clean = purchaseOrderNumber.substring(6).substring(0, 10);
                    }
                    if (result1 == true && purchaseOrderNumber.substring(6).substring(0, 1).equals("1")) {
                        purchaseOrderNumber_clean = "5" + purchaseOrderNumber.substring(6);
                        purchaseOrderNumber_clean = purchaseOrderNumber_clean.substring(0, 10);
                    }
                } else {
                    if (result2 == true && purchaseOrderNumber.substring(0, 2).equals("56")) {
                        purchaseOrderNumber_clean = purchaseOrderNumber.substring(0, 10);
                    } else if (result3 == true && purchaseOrderNumber.substring(0, 1).equals("1")) {
                        purchaseOrderNumber_clean = "5" + purchaseOrderNumber;
                        purchaseOrderNumber_clean = purchaseOrderNumber_clean.substring(0, 10);
                    } else {
                        Pattern p = Pattern.compile("[^0-9]");//提取有效数字
                        purchaseOrderNumber_clean = p.matcher(purchaseOrderNumber).replaceAll("").trim();
                    }
                    purchaseOrderNumber_clean = purchaseOrderNumber_clean.substring(0, 10);
                    StringBuilder sb = new StringBuilder(purchaseOrderNumber_clean);
                    sb.replace(0, 2, "56");
                    purchaseOrderNumber_clean = sb.toString();
                }
                purchaseOrderNumber_clean = purchaseOrderNumber_clean.replace("\n", "");
                Float float_parse = Float.parseFloat(purchaseOrderNumber_clean);
            } catch (Exception e) {
                purchaseOrderNumber_clean = purchaseOrderNumber + " ";
            }
        }
        return purchaseOrderNumber_clean;

    }

    /**
     * 清洗quantity业务
     *
     * @param quantity
     * @return
     */
    public static String cleanQuantity(String quantity) {
        if (!"".equals(quantity)) {
            quantity = quantity.replace(",", "");
            quantity = quantity.replace("，", "");
            quantity = quantity.replace(" ", "");
            quantity = quantity.replace("-", "");
            quantity = quantity.replace(".", "");
            quantity = quantity.replace("I", "1");
            quantity = quantity.replaceAll("(?:I|!)", "1");
            quantity = quantity.replace("(", "1");
            quantity = quantity.replace(")", "1");
            quantity = quantity.replace("（", "1");
            quantity = quantity.replace("）", "1");
            quantity = quantity.replace("&", "8");
            quantity = quantity.replace("o", "0");
            quantity = quantity.replace("O", "0");
            quantity = quantity.replace("Q", "0");
            quantity = quantity.replace("（", "1");
            quantity = quantity.replace("）", "1");
            quantity = quantity.replace("|", "1");
            quantity = quantity.replace("\\", "1");
            quantity = quantity.replace("/", "1");
            quantity = quantity.replace("\n", "");
            quantity = quantity.replace("↓", "1");
            quantity = quantity.replace("■", "");
            quantity = quantity.replace("•", "");
            quantity = quantity.replace("z", "");
            quantity = quantity.replace("M", "14");
            quantity = quantity.replace("-1", "4");
            quantity = quantity.replace("S", "8");
            quantity = quantity.replace("s", "8");
            quantity = quantity.trim();
            quantity = quantity.substring(0, quantity.length() - 3) + "." + quantity.substring(quantity.length() - 3);
            try {
                Float quantity_parse = Float.parseFloat(quantity);
            } catch (Exception e) {
                quantity = quantity + " ";
            }
        }
        return quantity;
    }

    /**
     * 清洗poShortText业务
     *
     * @param dataDO
     * @return
     */
    public static String cleanShortText(DataDO dataDO) {
        String poShortText = dataDO.getPoshorttext();
        if (!"".equals(poShortText)) {
            poShortText = poShortText.replace(" ", "");
            poShortText = poShortText.replace(",", "");
            poShortText = poShortText.replace("，", "");
            poShortText = poShortText.replace("'", "");
            poShortText = poShortText.replace("’", "");
            poShortText = poShortText.replace("”", "");
            poShortText = poShortText.replace("\"", "");
            poShortText = poShortText.replace("·", "");
            poShortText = poShortText.replace(".", "");
            poShortText = poShortText.replace("xI", "x1");
            poShortText = poShortText.replace("xl", "x1");
            poShortText = poShortText.replace("HSII", "HSU");
            poShortText = poShortText.replace("HSIJ", "HSU");
            poShortText = poShortText.replace("HSlJ", "HSU");
            poShortText = poShortText.replace("HS1J", "HSU");
            poShortText = poShortText.replace("HSiJ", "HSU");
            poShortText = poShortText.replace("HS|J", "HSU");
            poShortText = poShortText.replace("HS[I", "HSU");
            poShortText = poShortText.replace("HS[J", "HSU");
            poShortText = poShortText.replace("HSLI", "HSU");
            poShortText = poShortText.replace("HSL]", "HSU");
            poShortText = poShortText.replace("HSLl", "HSU");
            poShortText = poShortText.replace("HSL|", "HSU");
            poShortText = poShortText.replace("HSLJ", "HSU");
            poShortText = poShortText.replace("HSLi", "HSU");
            poShortText = poShortText.replace("HSL1", "HSU");
            poShortText = poShortText.replace("HSU-!", "HSU-1");
            poShortText = poShortText.replace("HRP", "HRF");
            poShortText = poShortText.replace("HRT", "HRF");
            StringBuffer stringBuffer = new StringBuffer(poShortText);
            String str = poShortText.substring(0, 3);
            if ("FIR".equals(str) || "l-I".equals(str) || "IIR".equals(str) || "liR".equals(str) || "TIR".equals(str)) {
                poShortText = "HR" + poShortText.substring(3);
            }
            if ("HR-".equals(str)) {
                poShortText = getString_HR(poShortText);
            }
            if ("HRF-".equals(poShortText.substring(0, 4))) {
                poShortText = getString_HRF(poShortText);
            }
            if ("HCF".equals(str)) {
                poShortText = poShortText.replace("WW", "");
                poShortText = poShortText.replace("\n", "");
            }
            if ("HWM".equals(str)) {
                if (poShortText.substring(poShortText.length() - 4).equals("2015")) {
                    stringBuffer.replace(stringBuffer.length() - 4, stringBuffer.length(), "");
                }
                if (poShortText.substring(poShortText.length() - 4).equals("CPP)")) {
                    stringBuffer.replace(stringBuffer.length() - 4, stringBuffer.length(), "(PP)");
                }
                poShortText = new String(stringBuffer);
            } else {
                poShortText = getString_HSU(poShortText);
            }
        } else {
            poShortText = "";
        }
        return poShortText;
    }

    public static String getString_HSU(String poShortText) {
        if (poShortText.substring(0, 3).equals("HSU")) {
            poShortText = poShortText.replace("O3T", "03T");
            poShortText = poShortText.replace("^^", "-");
            poShortText = poShortText.replace("\n", "");
            poShortText = poShortText.replace("⑻", "(N)");
            poShortText = poShortText.replace("CN", "(N");
            poShortText = poShortText.replace("CN1", "(N)");
            poShortText = poShortText.replace("-I", "-1");
            poShortText = poShortText.replace("-l", "-1");
            poShortText = poShortText.replace("-|", "-1");
            poShortText = poShortText.replace("-L", "-1");
            poShortText = poShortText.replace("VF0", "VFB");
            poShortText = poShortText.replace("(N1", "(N)");
            poShortText = poShortText.replace("O", "0");
        }
        return poShortText;
    }

    /**
     * 清洗以HR开头的数据
     *
     * @param poshorttext
     * @return
     */
    public static String getString_HR(String poshorttext) {
        System.out.println("输出：" + poshorttext);
        poshorttext = poshorttext.replace("CBQ18", "CEQ18");
        poshorttext = poshorttext.replace("CBQ15", "CEQ15");
        StringBuffer sb = new StringBuffer(poshorttext);
        if (poshorttext.substring(6, 7).equals("I")) {
            sb.replace(6, 7, "1");
        }
        if ("H0".equals(poshorttext.substring(8, 10))) {
            sb.replace(8, 10, "HO ");
        }
        sb.insert(sb.length() - 2, " ");
        return sb.toString();
    }

    /**
     * 清洗以HRF开头的数据
     *
     * @param pos
     * @return
     */
    public static String getString_HRF(String pos) {
        StringBuffer sb = new StringBuffer(pos);
        String str = pos.substring(pos.length() - 2);
        if ("OB".equals(str) || "0B".equals(str)) {
            sb.replace(pos.length() - 2, pos.length(), "GB");
        }
        if (!sb.toString().contains("54IMB")) {
            sb.insert(sb.length() - 2, " ");
        }
        return sb.toString();
    }

    /**
     * 生成SAP表
     *
     * @param dataDOS
     * @param filename
     */
    public static void excelOutput_SAP(List<DataDO> dataDOS, String filename) throws Exception {
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        XSSFSheet xssfSheet = xssfWorkbook.createSheet("Sheet1");
        String[] headers = new String[]{};
        headers = new String[]{"FileName", "CompanyCode", "Amount", "InvoiceDate", "InvoiceReferenceNumber", "InvoiceReferenceNumber2", "POShortText", "PurchaseOrderNumber", "Quantity", "TaxAmount", "TotalAmount", "UnitPrice", "Currency", "SONumber", "GoodsDescription", "OCRStatus", "PostingDate", "TaxCode", "Text", "BaselineDate", "ExchangeRate", "PaymentBlock", "Assignment", "HeaderText", "Item"};
        Row row0 = xssfSheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            XSSFCell cell = (XSSFCell) row0.createCell(i);
            XSSFRichTextString text = new XSSFRichTextString(headers[i]);
            cell.setCellValue(text);
        }
        int rowNum = 1;
        for (DataDO dataDO : dataDOS) {
            String poShortText = "";
            if ("CONN".equals(dataDO.getPoshorttext().substring(0, 4))) {
                poShortText = "PIPE";
            } else {
                poShortText = dataDO.getPoshorttext();
            }
            XSSFRow row1 = xssfSheet.createRow(rowNum);
            row1.createCell(0).setCellValue(dataDO.getFilepath());
            row1.createCell(1).setCellValue(dataDO.getCompanyCode());
            row1.createCell(2).setCellValue(dataDO.getAmount());
            row1.createCell(3).setCellValue(dataDO.getInvoicedate());
            row1.createCell(4).setCellValue(dataDO.getInvoiceReferenceNumber());
            row1.createCell(5).setCellValue(dataDO.getInvoiceReferenceNumber2());
            row1.createCell(6).setCellValue(poShortText);
            row1.createCell(7).setCellValue(dataDO.getPurchaseOrderNumber());
            row1.createCell(8).setCellValue(dataDO.getQuantity());
            row1.createCell(9).setCellValue(dataDO.getTaxAmount());
            row1.createCell(10).setCellValue(dataDO.getTotalAmount());
            row1.createCell(11).setCellValue(dataDO.getUnitPrice());
            row1.createCell(12).setCellValue("THB");
            row1.createCell(13).setCellValue(dataDO.getSOnumber());
            row1.createCell(14).setCellValue(dataDO.getGoodDescription());
            row1.createCell(15).setCellValue("OK");
            row1.createCell(16).setCellValue(dataDO.getPostingDate());
            row1.createCell(17).setCellValue(dataDO.getTaxCode());
            row1.createCell(18).setCellValue(dataDO.getPurchaseOrderNumber());
            row1.createCell(19).setCellValue(dataDO.getInvoicedate());
            row1.createCell(20).setCellValue("");
            row1.createCell(21).setCellValue("");
            row1.createCell(22).setCellValue(dataDO.getInvoiceReferenceNumber2());
            row1.createCell(23).setCellValue("");
            row1.createCell(24).setCellValue(dataDO.getItem());
            rowNum++;
        }
        FileOutputStream fileOutputStream = new FileOutputStream(filename);
        xssfWorkbook.write(fileOutputStream);
    }


    public static List getFilePathDate(String filePath) throws Exception {
        String invoicenum = "";
        String item = "";
        File excelFile = new File(filePath);
        FileInputStream excelData = new FileInputStream(excelFile);
        List<DataDO> dataDoList = new ArrayList<>();
        Workbook wb =  WorkbookFactory.create(excelData);
        Sheet sheet = wb.getSheet("inv");
        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row rows = sheet.getRow(r);//获取第r行
            if (rows.getCell(5) != null){
                rows.getCell(5).setCellType(Cell.CELL_TYPE_STRING);
                item = rows.getCell(5).getStringCellValue();
            }else {
                item = "";
            }
            if (rows.getCell(9) != null){
                rows.getCell(9).setCellType(Cell.CELL_TYPE_STRING);
                invoicenum = rows.getCell(9).getStringCellValue();
            }else {
                invoicenum = "";
            }
            DataDO dataDO = new DataDO();
            dataDO.setInvoicenum(invoicenum);
            dataDO.setItem(item);
            dataDoList.add(dataDO);
        }
        return dataDoList;
    }



}
