package HttpUtil;

import DataClean.BaseUtil;
import DataClean.DataDO;
import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.commons.io.FileUtils;

import java.io.File;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 德国，巴基斯坦，以及其他国家从新加坡采购发票数据处理
 */
public class Singapore_Hrois extends BaseUtil {
    public static final String QUANTITY_SGP = "SETS";
    public static final String REGER_SGP = "^([a-zA-Z]{3}+[0-9]+.+[0-9])+$";

    public static void main(String[] args) throws Exception {
        String companyCode = args[0];
        String filePath = args[1];
        String OCR_filePath = args[2];
        String ETL_filePath = args[3];


        //String companyCode="6280-HQ";
//        String filePath = "F:\\img\\德国\\4-0.jpg";
//        String OCR_filePath ="F:\\img\\德国\\OCR_file.xls";
//        String ETL_filePath= "F:\\img\\德国\\ETL_file.xls";
//
//        String companyCode = "6620-HQ";
//        String filePath = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AP\\新加坡用horis的国家\\巴基斯坦\\20190429\\6620-HQ-3812611INVPL-0.jpg";
//        String OCR_filePath = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AP\\新加坡用horis的国家\\巴基斯坦\\20190429\\OCR_file.xls";
//        String ETL_filePath = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AP\\新加坡用horis的国家\\巴基斯坦\\20190429\\ETL_file.xls";

//        String companyCode = "6550-HQ";
//        String filePath= "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AP\\新加坡用horis的国家\\马来西亚\\1-0-母版.jpg";
//        String OCR_filePath = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AP\\新加坡用horis的国家\\马来西亚\\test\\OCR_file.xls";
//        String ETL_filePath= "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AP\\新加坡用horis的国家\\马来西亚\\test\\ETL_file.xls";
//
//        String companyCode = "6400-HQ";
//        String filePath = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AP\\新加坡用horis的国家\\日本\\test\\201904161845\\45007418774500742893-inv.jpg";
//        String OCR_filePath = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AP\\新加坡用horis的国家\\日本\\test\\201904161845\\OCR_file.xls";
//        String ETL_filePath = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AP\\新加坡用horis的国家\\日本\\test\\201904161845\\ETL_file.xls";

//        String companyCode = "62H0-HQ";
//        String filePath= "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AP\\新加坡用horis的国家\\法国_比利时\\比利时\\62H0-HQ-3-0\\pic\\1.jpg\\0003748989INV_1.jpg";
//        String OCR_filePath = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AP\\新加坡用horis的国家\\法国_比利时\\比利时\\62H0-HQ-3-0\\pic\\1.jpg\\OCR_file.xls";
//        String ETL_filePath= "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AP\\新加坡用horis的国家\\法国_比利时\\比利时\\62H0-HQ-3-0\\pic\\1.jpg\\ETL_file.xls";
        //从CompanyCode 中截取 PONUmber

        String fileName = filePath.substring(filePath.lastIndexOf("\\") + 1);//截取图片路径
        //不同的companyCode需要不同的OCR模板
        try {
            switch (companyCode) {
                case "6280-HQ"://德国发票数据处理
                    clearDEFunction(companyCode, filePath, OCR_filePath, ETL_filePath, fileName);
                    break;
                case "6620-HQ"://巴基斯坦发票数据处理
                    System.out.println("走的巴基斯坦");
                    clearPAKFunction(companyCode, filePath, OCR_filePath, ETL_filePath, fileName);
                    break;
                case "6550-HQ"://马来西亚发票数据处理

                    clearMASFunction(companyCode, filePath, OCR_filePath, ETL_filePath, fileName);
                    break;
                case "65G0-HQ"://菲律宾发票数据处理
                    clearPHIFunction(companyCode, filePath, OCR_filePath, ETL_filePath, fileName);
                    break;
                case "62G0-HQ"://法国
                case "62H0-HQ"://比利时
                    clearFAR_BELFunction(companyCode, filePath, OCR_filePath, ETL_filePath, fileName);
                    break;
                default://其它国家发票数据处理

                    clearOther_SGPFunction(companyCode, filePath, OCR_filePath, ETL_filePath, fileName);
                    break;
            }
            System.out.println("运行结束！！！");
        } catch (Exception e) {
            String msg = "OCR Data Incompleted";
            templateNot_SGP(companyCode, fileName, ETL_filePath, OCR_filePath, msg);
        }
    }

    /**
     * 德国发票数据处理
     *
     * @param companyCode
     * @param filePath
     * @param OCR_filePath
     * @param ETL_filePath
     * @param fileName
     * @throws Exception
     */
    public static void clearDEFunction(String companyCode, String filePath, String OCR_filePath, String ETL_filePath, String fileName) throws Exception {
        String invoiceReferenceNumber = "";
        String totalAmount = "";
        String currency = "";
        String invoiceDate = "";
        String purchaseOrderNumber = "";
        File file = new File(filePath);
        byte[] fileData = FileUtils.readFileToByteArray(file);//读取图片
        //OCR模板扫描 返回结果数据
        String json = ocrImageFile("APSigGer", file.getName(), fileData);
        System.out.println("输出德国json:" + json);
        JSONObject jsonObject = JSON.parseObject(json);
        //判断是否返回扫描成功数据
        if (jsonObject.get("result").toString().contains("success")) {
            //处理json数据
            JSONArray jsonArray = jsonObject.getJSONObject("ocrResult").getJSONArray("ranges");
            Map<String, Object> map = new HashMap<>();
            for (int i = 0; i < jsonArray.size(); i++) {
                JSONObject object = (JSONObject) jsonArray.get(i);
                switch (object.get("rangeId").toString()) {
                    case "InvoiceDate":
                        if (object.get("value") != null) {
                            invoiceDate = clearInvoiceDate(object.get("value").toString());
                        }
                        break;
                    case "PurchaseOrderNumber":
                        if (object.get("value") != null) {
                            purchaseOrderNumber = checkPONumber(object.get("value").toString());
                        }
                        break;
                    case "InvoiceReferenceNumber":
                        if (object.get("value") != null) {
                            invoiceReferenceNumber = clearIRNumber(object.get("value").toString());
                        }
                        break;
                    case "POShortText":
                        if (object.get("value") != null) {
                            //通过空格拆分，获取每一个数据
                            String[] vl = object.get("value").toString().split(" ");
                            List<String> list = Arrays.asList(vl);
                            map.put("POShortText", list);
                        } else {
                            List<String> list = new ArrayList<>();
                            map.put("POShortText", list);
                        }
                        break;
                    case "Quantity":
                        if (object.get("value") != null) {
                            String[] vl = object.get("value").toString().split(" ");
                            List<String> list1 = new ArrayList<>();
                            List<String> list = Arrays.asList(vl);
                            for (String quantity : list) {
                                Pattern p = Pattern.compile(REGER);//提取有效数字
                                quantity = p.matcher(quantity).replaceAll("").trim();
                                if (!"".equals(quantity)) {
                                    list1.add(quantity);
                                }
                            }
                            map.put("Quantity", list1);
                        } else {
                            List<String> list = new ArrayList<>();
                            map.put("Quantity", list);
                        }
                        break;
                    case "Price":
                        if (object.get("value") != null) {
                            //通过空格无法很好的拆分所有数据，需要将数据中的空格去掉，通过货币号来进行拆分数据
                            String unitPrice = object.get("value").toString().replace(" ", "");
                            String[] regerSplit = REGER_CURRENCY;//获取货币号数组
                            List<String> list1 = new ArrayList<>();
                            for (String curr : regerSplit) {
                                if (unitPrice.contains(curr)) {
                                    String[] vl = unitPrice.split(curr);
                                    List<String> list = Arrays.asList(vl);
                                    for (int j = 1; j < list.size(); j++) {
                                        String price = clearData(list.get(j));
                                        if (!"".equals(price)) {
                                            list1.add(price);
                                        }
                                    }
                                    currency = curr;//获取货币号
                                    break;
                                }
                            }
                            map.put("UnitPrice", list1);
                        } else {
                            List<String> list = new ArrayList<>();
                            map.put("UnitPrice", list);
                        }
                        break;
                    //OCR扫描中，amount以及totalAmount同时获得，
                    case "Amount":
                        if (object.get("value") != null) {
                            //通过空格无法很好的拆分所有数据，需要将数据中的空格去掉，通过货币号来进行拆分数据
                            String amountAndTotal = object.get("value").toString().replace(" ", "");
                            String[] regerSplit = REGER_CURRENCY;//获取货币号数组
                            List<Double> list1 = new ArrayList<>();
                            for (String curr : regerSplit) {
                                if (amountAndTotal.contains(curr)) {
                                    String[] vl = amountAndTotal.split(curr);
                                    List<String> list = Arrays.asList(vl);
                                    for (int j = 1; j < list.size(); j++) {
                                        String amount = clearData(list.get(j));
                                        if (!"".equals(amount)) {
                                            list1.add(Double.parseDouble(amount));
                                        }
                                    }
                                    break;
                                }
                            }
                            map.put("Amount", list1);
                        } else {
                            List<String> list = new ArrayList<>();
                            map.put("Amount", list);
                        }
                        break;
                }
            }
            System.out.println("输出currency：" + currency);
            //在数据获取是，totalAmount的值会痛amount值一起扫描得到
            //获取amount的集合，获取集合中做大的值即为totalamount，同时将最大值从集合中删除
            List<Double> amountList = (List<Double>) map.get("Amount");
            if (amountList.size() > 0) {
                totalAmount = String.valueOf(Collections.max(amountList));//获取totalAmount
                int j = amountList.lastIndexOf(Double.parseDouble(totalAmount));
                amountList.remove(j);
            }
            for (String key : map.keySet()) {//查看
                System.out.println(key + ":" + map.get(key));
            }
            List<String> poShortList = (List<String>) map.get("UnitPrice");
            List<String> priceList = (List<String>) map.get("UnitPrice");
            List<String> quantityList = (List<String>) map.get("Quantity");
            //判断所提取获得的数据条目数是否数量相等、不相等即为OCR模板不适用
            if (quantityList.size() > 0 && quantityList.size() == priceList.size() && quantityList.size() == amountList.size() && quantityList.size() == poShortList.size()) {
                List<DataDO> dataDOList = new ArrayList<>();
                for (int i = 0; i < quantityList.size(); i++) {
                    DataDO dataDO = new DataDO();
                    dataDO.setDownloadStatus("OK");
                    dataDO.setCompanyCode(companyCode);
                    dataDO.setFilepath(fileName);
                    dataDO.setInvoicedate(invoiceDate);
                    dataDO.setPurchaseOrderNumber(purchaseOrderNumber);
                    dataDO.setPoshorttext(poShortList.get(i));
                    dataDO.setAmount(convertData(String.valueOf(amountList.get(i))));
                    dataDO.setInvoiceReferenceNumber(invoiceReferenceNumber);
                    dataDO.setQuantity(quantityList.get(i));
                    dataDO.setTotalAmount(convertData(totalAmount));
                    dataDO.setUnitPrice(priceList.get(i));
                    dataDO.setOCRStatus("OK");
                    dataDO.setCurrency(currency);
                    dataDOList.add(dataDO);
                }
                //生成ETL表
                excelOutput_OCR(dataDOList, ETL_filePath);
                //生成OCR表
                excelOutput_OCR(dataDOList, OCR_filePath);
            } else {
                String msg = "OCR Data Incompleted";
                templateNot_SGP(companyCode, fileName, ETL_filePath, OCR_filePath, msg);
            }
        } else {
            String msg = "OCR Template not match";
            templateNot_SGP(companyCode, fileName, ETL_filePath, OCR_filePath, msg);
        }
    }

    /**
     * 其它国家发票数据处理
     *
     * @param companyCode
     * @param filePath
     * @param OCR_filePath
     * @param ETL_filePath
     * @param fileName
     * @throws Exception
     */
    public static void clearOther_SGPFunction(String companyCode, String filePath, String OCR_filePath, String ETL_filePath, String fileName) throws Exception {
        File file = new File(filePath);
        byte[] fileData = FileUtils.readFileToByteArray(file);//读取图片
        //使用第一个OCR模板扫描 返回结果数据
        String json = ocrImageFile("APSigUnified_0109", file.getName(), fileData);
        JSONObject jsonObject = JSON.parseObject(json);
        System.out.println("输出新加坡json:" + json);
        //判断结果数据返回情况
        if (jsonObject.get("result").toString().contains("success")) {
            clearAPSigUnifiedOne(jsonObject, companyCode, filePath, OCR_filePath, ETL_filePath, fileName);
        } else {
            String msg = "OCR Template not match";
            templateNot_SGP(companyCode, fileName, ETL_filePath, OCR_filePath, msg);
        }
    }

    /**
     * 过滤data(可以直接使用空格将每条数据拆分出来时使用)
     *
     * @param data
     * @return
     */
    public static String checkData(String data) {
        data = data.replace("O", "0");
        data = data.replace("i", "1");
        data = data.replace("l", "1");
        data = data.replace("Q", "0");
        data = data.replace("C", "G");
        Pattern p = Pattern.compile("[\u4e00-\u9fa5]");
        Matcher m = p.matcher(data);
        if (m.find()) {
            data = "";
        }
        return data;
    }

    /**
     * 过滤POSONumber
     *
     * @param data
     * @return
     */
    public static String checkPOSO(String data) {
        if (data != null && !"".equals(data)) {
            data = data.replace(" ", "");
            data = data.replace("4S", "45");
            data = data.replace("N0", "");
            data = data.replace("P/0", "");
            data = data.replace("S/0", "");
            data = data.substring(data.indexOf("4500", 0));
        }
        return data;
    }

    /**
     * 过滤POShortText
     *
     * @param data
     * @return
     */
    public static String checkPOShortText(String data) {
        if (data != null && !"".equals(data)) {
            data = data.replace("CUK1)", "(UK1)");
            data = data.replace("UI<", "UK");
            data = data.replace("CUI<)", "(UK)");
            data = data.replace("CW)", "CWJ");
            data = data.replace("Z", "2");
            data = data.replace("OM", "0M");
        }
        return data;
    }

    /**
     * 新加坡采购发票数据的模板不适用
     *
     * @param companyCode
     * @param fileName
     * @param ETL_filePath
     * @param OCR_filePath
     */
    public static void templateNot_SGP(String companyCode, String fileName, String ETL_filePath, String OCR_filePath, String message) {
        List<DataDO> dataDOList = new ArrayList<>();
        DataDO dataDO = new DataDO();
        dataDO.setCompanyCode(companyCode);
        dataDO.setFilepath(fileName);
        dataDO.setOCRStatus(message);
        dataDO.setDownloadStatus("OK");
        dataDOList.add(dataDO);
        //生成ETL表
        excelOutput_OCR(dataDOList, ETL_filePath);
        //生成OCR表
        excelOutput_OCR(dataDOList, OCR_filePath);
    }

    /**
     * 清洗poNumber
     *
     * @param posoNumber
     * @return
     */
    public static String checkPONumber(String posoNumber) {
        Pattern p = Pattern.compile("[^0-9]");//提取有效数字
        posoNumber = p.matcher(posoNumber).replaceAll("").trim();
        List<String> list = new ArrayList<>();
        for (int k = 0; k < posoNumber.length(); k += 10) {
            String str = "";
            if (k + 10 > posoNumber.length()) {
                str = posoNumber.substring(k);
            } else {
                str = posoNumber.substring(k, k + 10);
            }
            list.add(str);
        }
        String pNumber = "";
        for (String pos : list) {
            if (pos.indexOf("4500", 0) == 0 && pos.length() == 10) {
                pNumber = pNumber + pos + " ";
            }
        }
        posoNumber = pNumber;
        System.out.println("po号：" + posoNumber);
        return posoNumber;
    }

    /**
     * 巴基斯坦发票数据处理
     * 巴基斯坦发票有两种类型，所以在模板扫描获取结果时判断使用哪个模板
     *
     * @param companyCode
     * @param filePath
     * @param OCR_filePath
     * @param ETL_filePath
     * @param fileName
     * @throws Exception
     */
    public static void clearPAKFunction(String companyCode, String filePath, String OCR_filePath, String ETL_filePath, String fileName) throws Exception {
        File file = new File(filePath);
        byte[] fileData = FileUtils.readFileToByteArray(file);//读取图片
//        //使用第一个OCR模板扫描 返回结果数据
//        String json = ocrImageFile("APSigPak01", file.getName(), fileData);
//        JSONObject jsonObject = JSON.parseObject(json);
        //判断结果数据返回情况
//        if (jsonObject.get("result").toString().contains("success")) {
//            System.out.println("输出巴基斯坦json:"+ json);
//            clearPAKTemplateOne(jsonObject, companyCode,PONumber, filePath, OCR_filePath, ETL_filePath, fileName);
//        }else {
        //若模板以返回结果不成功，则使用第二个模板进行扫描
        String json2 = ocrImageFile("APSigPak02", file.getName(), fileData);
        System.out.println("输出巴基斯坦json2:" + json2);
        JSONObject jsonObject2 = JSON.parseObject(json2);
        if (jsonObject2.get("result").toString().contains("success")) {
            clearPAKTemplateTwo(jsonObject2, companyCode, filePath, OCR_filePath, ETL_filePath, fileName);
        } else {
            String msg = "OCR Template not match";
            templateNot_SGP(companyCode, fileName, ETL_filePath, OCR_filePath, msg);
        }
//        }
    }

    /**
     * 巴基斯坦模板一数据处理
     *
     * @param jsonObject
     */
    public static void clearPAKTemplateOne(JSONObject jsonObject, String companyCode, String PONumber, String filePath, String OCR_filePath, String ETL_filePath, String fileName) {
        String invoiceReferenceNumber = "";
        String totalAmount = "";
        String currency = "";
        String invoiceDate = "";
        String posoNumber = "";
        JSONArray jsonArray = jsonObject.getJSONObject("ocrResult").getJSONArray("ranges");
        Map<String, Object> map = new HashMap<>();
        for (int i = 0; i < jsonArray.size(); i++) {
            JSONObject object = (JSONObject) jsonArray.get(i);
            switch (object.get("rangeId").toString()) {
                case "InvoiceDate":
                    if (object.get("value") != null) {
                        invoiceDate = clearInvoiceDate(object.get("value").toString());
                    }
                    break;
//                case "POSONumber":
//                    if (object.get("value") != null){
//                        posoNumber = object.get("value").toString();
//                        posoNumber = posoNumber.substring(posoNumber.indexOf("6200"),posoNumber.indexOf("6200")+10);
//                    }
//                    break;
                case "POShortText":
                    if (object.get("value") != null) {
                        String[] vl = object.get("value").toString().split(" ");
                        List<String> list = Arrays.asList(vl);
                        map.put("POShortText", list);
                    } else {
                        List<String> list = new ArrayList<>();
                        map.put("POShortText", list);
                    }
                    break;
                case "Quantity":
                    if (object.get("value") != null) {
                        //通过空格拆分数据
                        String[] vl = object.get("value").toString().split(" ");
                        List<String> list = Arrays.asList(vl);
                        List<String> list1 = new ArrayList<>();
                        for (String quantity : list) {
                            quantity = checkData(quantity);
                            if (quantity.contains(QUANTITY_SGP)) {
                                Pattern p = Pattern.compile(REGER);//提取有效数字
                                quantity = p.matcher(quantity).replaceAll("").trim();
                                if (!"".equals(quantity)) {
                                    list1.add(quantity);
                                }
                            }
                        }
                        map.put("Quantity", list1);
                    } else {
                        List<String> list = new ArrayList<>();
                        map.put("Quantity", list);
                    }
                    break;
                case "InvoiceReferenceNumber":
                    if (object.get("value") != null) {
                        invoiceReferenceNumber = clearIRNumber(object.get("value").toString());
                    } else {
                        invoiceReferenceNumber = "";
                    }
                    break;
                case "Price":
                    if (object.get("value") != null) {
                        //通过空格拆分数据
                        String[] vl = object.get("value").toString().split(" ");
                        List<String> list = Arrays.asList(vl);
                        List<String> list1 = new ArrayList<>();
                        for (String price : list) {
                            price = checkData(price);
                            if (price.matches(REGER_SGP)) {//获取满足条件的数据
                                list1.add(price);
                            }
                        }
                        //获取currency
                        Pattern p = Pattern.compile("[a-zA-Z]");
                        Matcher m = p.matcher(list1.get(0));
                        while (m.find()) {
                            currency = currency + m.group(0);
                        }
                        map.put("UnitPrice", list1);
                    } else {
                        List<String> list = new ArrayList<>();
                        map.put("UnitPrice", list);
                    }
                    break;
                case "Amount":
                    if (object.get("value") != null) {
                        //通过空格拆分数据
                        String[] vl = object.get("value").toString().split(" ");
                        List<String> list = Arrays.asList(vl);
                        List<Double> list1 = new ArrayList<>();
                        for (String amount : list) {
                            amount = checkData(amount);
                            if (amount.matches(REGER_SGP)) {
                                amount = amount.replaceAll("[a-zA-Z]", "").trim();
                                list1.add(Double.parseDouble(amount));
                            }
                        }
                        map.put("Amount", list1);
                    } else {
                        List<String> list = new ArrayList<>();
                        map.put("Amount", list);
                    }
                    break;
            }
        }
        System.out.println("输出currency：" + currency);
        //获取amount集合，获取集合的最大值即为totalAmount，并将最大值删除
        List<Double> amountList = (List<Double>) map.get("Amount");
        if (amountList.size() > 0) {
            totalAmount = String.valueOf(Collections.max(amountList));//获取最大值
            int i = amountList.lastIndexOf(Double.parseDouble(totalAmount));//获取最大值最后出现的位置
            amountList.remove(i);//删除
        }
        for (String key : map.keySet()) {//查看
            System.out.println(key + ":" + map.get(key));
        }
        List<String> poShortTextList = (List<String>) map.get("POShortText");
        List<String> priceList = (List<String>) map.get("UnitPrice");
        List<String> quantityList = (List<String>) map.get("Quantity");
        //判断获取的数据数量是否相等，不相等即为OCR模板不适用
        if (quantityList.size() == poShortTextList.size() && quantityList.size() == priceList.size() && quantityList.size() == amountList.size() && quantityList.size() > 0) {
            List<DataDO> dataDOList = new ArrayList<>();
            for (int i = 0; i < quantityList.size(); i++) {
                DataDO dataDO = new DataDO();
                dataDO.setDownloadStatus("OK");
                dataDO.setCompanyCode(companyCode);
                dataDO.setInvoicedate(invoiceDate);
                dataDO.setPurchaseOrderNumber(PONumber);
                dataDO.setPoshorttext(poShortTextList.get(i));
                dataDO.setFilepath(fileName);
                dataDO.setAmount(convertData(String.valueOf(amountList.get(i))));
                dataDO.setInvoiceReferenceNumber(invoiceReferenceNumber);
                dataDO.setQuantity(quantityList.get(i));
                dataDO.setTotalAmount(convertData(totalAmount));
                dataDO.setUnitPrice(priceList.get(i).replaceAll("[a-zA-Z]", "").trim());
                dataDO.setCurrency(currency);
                dataDO.setOCRStatus("OK");
                dataDOList.add(dataDO);
            }
            //生成ETL表
            excelOutput_OCR(dataDOList, ETL_filePath);
            //生成OCR表
            excelOutput_OCR(dataDOList, OCR_filePath);
        } else {
            String msg = "OCR Data Incompleted";
            templateNot_SGP(companyCode, fileName, ETL_filePath, OCR_filePath, msg);
        }
    }

    /**
     * 巴基斯坦模板二数据处理
     *
     * @param jsonObject
     */
    public static void clearPAKTemplateTwo(JSONObject jsonObject, String companyCode,String filePath, String OCR_filePath, String ETL_filePath, String fileName) {
        String invoiceReferenceNumber = "";
        String totalAmount = "";
        String currency = "";
        String posoNumber = "";
        String invoiceDate = "";
        JSONArray jsonArray = jsonObject.getJSONObject("ocrResult").getJSONArray("ranges");
        Map<String, Object> map = new HashMap<>();
        for (int i = 0; i < jsonArray.size(); i++) {
            JSONObject object = (JSONObject) jsonArray.get(i);
            switch (object.get("rangeId").toString()) {
                case "InvoiceDate":
                    if (object.get("value") != null) {
                        invoiceDate = clearInvoiceDate(object.get("value").toString());
                    }
                    break;
                case "POSONumber":
                    if (object.get("value") != null){
                        posoNumber = object.get("value").toString();
                        if (posoNumber.contains("5200")){
                            posoNumber = posoNumber.substring(posoNumber.indexOf("5200"),posoNumber.indexOf("5200")+10);
                            posoNumber = posoNumber.replace("5200","6200");
                        }else if (posoNumber.contains("6200")){
                            posoNumber = posoNumber.substring(posoNumber.indexOf("6200"),posoNumber.indexOf("6200")+10);
                        }

                    }
                    break;
                case "POShortText":
                    if (object.get("value") != null) {
                        System.out.println("原始的POShortText:" + object.get("value").toString());
                        String[] vl = object.get("value").toString().split(" ");
                        List<String> list = new ArrayList<String>(Arrays.asList(vl));
                        if (list.get(0).length() < 3) {
                            list.remove(0);
                        }
                        map.put("POShortText", list);
                    } else {
                        List<String> list = new ArrayList<>();
                        map.put("POShortText", list);
                    }
                    break;
                case "Quantity":
                    if (object.get("value") != null) {
                        List<String> list1 = new ArrayList<>();
                        String[] vl = object.get("value").toString().split(" ");
                        List<String> list = Arrays.asList(vl);
                        for (String quantity : list) {
                            Pattern p = Pattern.compile(REGER);//提取有效数字
                            quantity = p.matcher(quantity).replaceAll("").trim();
                            if (!"".equals(quantity)) {
                                list1.add(quantity);
                            }
                        }
                        map.put("Quantity", list1);
                    } else {
                        List<String> list = new ArrayList<>();
                        map.put("Quantity", list);
                    }
                    break;
                case "InvoiceReferenceNumber":
                    if (object.get("value") != null) {
                        invoiceReferenceNumber = clearIRNumber(object.get("value").toString());
                    } else {
                        invoiceReferenceNumber = "";
                    }
                    break;
                case "Price":
                    if (object.get("value") != null) {
                        String unitPrice = object.get("value").toString().replace(" ", "");
                        System.out.println("输出price：" + unitPrice);
                        String[] regerSplit = REGER_CURRENCY;//获取货币号数组
                        List<String> list1 = new ArrayList<>();
                        for (String curr : regerSplit) {
                            if (unitPrice.contains(curr)) {
                                String[] vl = unitPrice.split(curr);
                                List<String> list = Arrays.asList(vl);
                                for (int j = 1; j < list.size(); j++) {
                                    String price = clearData(list.get(j));
                                    if (!"".equals(price)) {
                                        list1.add(price);
                                    }
                                }
                                break;
                            }
                        }
                        map.put("UnitPrice", list1);
                    } else {
                        List<String> list = new ArrayList<>();
                        map.put("UnitPrice", list);
                    }
                    break;
                //模板扫描时，amount以及totalAmount同时获得，
                case "Amount":
                    if (object.get("value") != null) {
                        String amountAndTotal = object.get("value").toString().replace(" ", "");
                        System.out.println("输出amountAndTotal：" + amountAndTotal);
                        List<Double> list1 = new ArrayList<>();
                        String[] regerSplit = REGER_CURRENCY;//获取货币号数组
                        for (String curr : regerSplit) {
                            if (amountAndTotal.contains(curr)) {
                                String[] vl = amountAndTotal.split(curr);
                                List<String> list = Arrays.asList(vl);
                                for (int j = 1; j < list.size(); j++) {
                                    String amount = clearData(list.get(j));
                                    if (!"".equals(amount)) {
                                        list1.add(Double.parseDouble(amount));
                                    }
                                }
                                currency = curr;
                                break;
                            }
                        }
                        map.put("Amount", list1);
                    } else {
                        List<String> list = new ArrayList<>();
                        map.put("Amount", list);
                    }
                    break;
            }
        }
        System.out.println("输出currency：" + currency);
        //获取amount集合，获取最大值即为totalAmount，并将最大值删除
        List<Double> amountList = (List<Double>) map.get("Amount");
        if (amountList.size() > 0) {
            totalAmount = String.valueOf(Collections.max(amountList));
            int i = amountList.lastIndexOf(Double.parseDouble(totalAmount));
            amountList.remove(i);
        }
        for (String key : map.keySet()) {
            System.out.println(key + ":" + map.get(key));
        }
        List<String> POShortTextList = (List<String>) map.get("POShortText");
        List<String> priceList = (List<String>) map.get("UnitPrice");
        List<String> quantityList = (List<String>) map.get("Quantity");
        System.out.println("quantityList:" + quantityList.size() + ",POShortTextList:" + POShortTextList.size() + ",priceList:" + priceList.size() + ",amountList:" + amountList.size());
        //判断获取的数据数量是否相等，不相等即为模板不适应
        if (quantityList.size() > 0 && quantityList.size() == POShortTextList.size() && quantityList.size() == priceList.size() && quantityList.size() == amountList.size()) {
            List<DataDO> dataDOList = new ArrayList<>();
            for (int i = 0; i < quantityList.size(); i++) {
                DataDO dataDO = new DataDO();
                dataDO.setDownloadStatus("OK");
                dataDO.setCompanyCode(companyCode);
                dataDO.setFilepath(fileName);
                dataDO.setInvoicedate(invoiceDate);
                dataDO.setPurchaseOrderNumber(posoNumber);
                dataDO.setPoshorttext(POShortTextList.get(i));
                dataDO.setAmount(convertData(String.valueOf(amountList.get(i))));
                dataDO.setInvoiceReferenceNumber(invoiceReferenceNumber);
                dataDO.setQuantity(quantityList.get(i));
                dataDO.setTotalAmount(convertData(totalAmount));
                dataDO.setUnitPrice(priceList.get(i));
                dataDO.setCurrency(currency);
                dataDO.setOCRStatus("OK");
                dataDOList.add(dataDO);
            }
            //生成ETL表
            excelOutput_OCR(dataDOList, ETL_filePath);
            //生成OCR表
            excelOutput_OCR(dataDOList, OCR_filePath);
        } else {
            String msg = "OCR Data Incompleted";
            templateNot_SGP(companyCode, fileName, ETL_filePath, OCR_filePath, msg);
        }
    }

    /**
     * 处理invoiceDate
     *
     * @param invoiceDate
     * @return
     */
    public static String clearInvoiceDate(String invoiceDate) {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
        SimpleDateFormat sdf1 = new SimpleDateFormat("MM/dd/yyyy");
        try {
            if (!"".equals(invoiceDate)) {
                invoiceDate = invoiceDate.replace("-", "/");
                invoiceDate = invoiceDate.replace(" ", "/");
                invoiceDate = sdf1.format(sdf.parse(invoiceDate));
            }
        } catch (ParseException e) {
            invoiceDate = invoiceDate + " ";
        }
        System.out.println("输出时间：" + invoiceDate);
        return invoiceDate;
    }

    /**
     * 马来西亚发票数据处理
     *
     * @param companyCode
     * @param filePath
     * @param OCR_filePath
     * @param ETL_filePath
     * @param fileName
     * @throws Exception
     */
    public static void clearMASFunction(String companyCode, String filePath, String OCR_filePath, String ETL_filePath, String fileName) throws Exception {
        String invoiceReferenceNumber = "";
        String totalAmount = "";
        String currency = "";
        String invoiceDate = "";
        String posoNumber = "";
        File file = new File(filePath);
        byte[] fileData = FileUtils.readFileToByteArray(file);//读取图片
        //OCR模板扫描 返回结果数据
        String json = ocrImageFile("APSigMal", file.getName(), fileData);
        System.out.println("输出马来西亚json:" + json);
        JSONObject jsonObject = JSON.parseObject(json);
        if (jsonObject.get("result").toString().contains("success")) {
            //处理json数据
            JSONArray jsonArray = jsonObject.getJSONObject("ocrResult").getJSONArray("ranges");
            Map<String, Object> map = new HashMap<>();
            for (int i = 0; i < jsonArray.size(); i++) {
                JSONObject object = (JSONObject) jsonArray.get(i);
                switch (object.get("rangeId").toString()) {
                    case "InvoiceDate":
                        if (object.get("value") != null) {
                            invoiceDate = clearInvoiceDate(object.get("value").toString());
                        }
                        break;
                    case "PurchaseOrderNumber":
                        if (object.get("value") != null) {
                            posoNumber = object.get("value").toString().replace("⑻", "00");
                        }
                        break;
                    case "InvoiceReferenceNumber":
                        if (object.get("value") != null) {
                            invoiceReferenceNumber = object.get("value").toString();
                        }
                        break;
                    case "POShortText":
                        if (object.get("value") != null) {
                            String[] value = object.get("value").toString().split(" ");
                            List<String> list = Arrays.asList(value);
                            List<String> list1 = new ArrayList<>();
                            for (String str : list) {
                                if (str.length() >= 7) {
                                    list1.add(str);
                                }
                            }
                            map.put("POShortText", list1);
                        } else {
                            List<String> list = new ArrayList<>();
                            map.put("POShortText", list);
                        }
                        break;
                    case "Quantity":
                        if (object.get("value") != null) {
                            String[] value = object.get("value").toString().split(" ");
                            List<String> list = Arrays.asList(value);
                            List<String> list1 = new ArrayList<>();
                            for (String quantity : list) {
                                if (quantity.matches("^([0-9]+[a-zA-Z]{3})+$")) {
                                    quantity = quantity.replaceAll("[a-zA-Z]", "").trim();
                                    list1.add(quantity);
                                }
                            }
                            map.put("Quantity", list1);
                        } else {
                            List<String> list = new ArrayList<>();
                            map.put("Quantity", list);
                        }
                        break;
                    case "Price":
                        if (object.get("value") != null) {
                            String value = object.get("value").toString().replace(" ", "");
                            List<String> list1 = new ArrayList<>();
                            String[] regerSplit = REGER_CURRENCY;//获取货币号数组
                            for (String curr : regerSplit) {
                                if (value.contains(curr)) {
                                    String[] vl = value.split(curr);
                                    List<String> list = Arrays.asList(vl);
                                    for (int j = 1; j < list.size(); j++) {
                                        String price = clearData(list.get(j));
                                        if (!"".equals(price)) {
                                            list1.add(price);
                                        }
                                    }
                                    break;
                                }
                            }
                            map.put("UnitPrice", list1);
                        } else {
                            List<String> list = new ArrayList<>();
                            map.put("UnitPrice", list);
                        }
                        break;
                    case "Amount":
                        if (object.get("value") != null) {
                            String value = object.get("value").toString().replace(" ", "");
                            List<Double> list1 = new ArrayList<>();
                            String[] regerSplit = REGER_CURRENCY;//获取货币号数组
                            for (String curr : regerSplit) {
                                if (value.contains(curr)) {
                                    String[] vl = value.split(curr);
                                    List<String> list = Arrays.asList(vl);
                                    for (int j = 1; j < list.size(); j++) {
                                        String amount = clearData(list.get(j));
                                        if (!"".equals(amount)) {
                                            list1.add(Double.parseDouble(amount));
                                        }
                                    }
                                    currency = curr;
                                    break;
                                }
                            }
                            map.put("Amount", list1);
                        } else {
                            List<String> list = new ArrayList<>();
                            map.put("Amount", list);
                        }
                        break;
                }
            }

            List<Double> amountList = (List<Double>) map.get("Amount");
            if (amountList.size() > 0) {
                System.out.println("输出最大值：" + Collections.max(amountList));
                totalAmount = String.valueOf(Collections.max(amountList));
                int j = amountList.lastIndexOf(Double.parseDouble(totalAmount));
                amountList.remove(j);
                map.put("Amount", amountList);
            }
            for (String key : map.keySet()) {//查看
                System.out.println(key + ":" + map.get(key));
            }
            List<String> priceList = (List<String>) map.get("UnitPrice");
            List<String> quantityList = (List<String>) map.get("Quantity");
            List<String> POShortTextList = (List<String>) map.get("POShortText");
            //判断提取获得的数据数量是否相等，不相等即为模板不适用
            if (quantityList.size() > 0 && quantityList.size() == priceList.size() && quantityList.size() == amountList.size() && quantityList.size() == POShortTextList.size()) {
                List<DataDO> dataDOList = new ArrayList<>();
                for (int i = 0; i < quantityList.size(); i++) {
                    DataDO dataDO = new DataDO();
                    dataDO.setDownloadStatus("OK");
                    dataDO.setCompanyCode(companyCode);
                    dataDO.setFilepath(fileName);
                    dataDO.setInvoicedate(invoiceDate);
                    dataDO.setAmount(convertData(String.valueOf(amountList.get(i))));
                    dataDO.setInvoiceReferenceNumber(invoiceReferenceNumber);
                    dataDO.setPurchaseOrderNumber(posoNumber);
                    dataDO.setPoshorttext(POShortTextList.get(i));
                    dataDO.setQuantity(quantityList.get(i));
                    dataDO.setTotalAmount(convertData(totalAmount));
                    dataDO.setUnitPrice(priceList.get(i));
                    dataDO.setOCRStatus("OK");
                    dataDO.setCurrency(currency);
                    dataDOList.add(dataDO);
                }
                //生成ETL表
                excelOutput_OCR(dataDOList, ETL_filePath);
                //生成OCR表
                excelOutput_OCR(dataDOList, OCR_filePath);
            } else {
                String msg = "OCR Data Incompleted";
                templateNot_SGP(companyCode, fileName, ETL_filePath, OCR_filePath, msg);
            }
        } else {
            String msg = "OCR Template not match";
            templateNot_SGP(companyCode, fileName, ETL_filePath, OCR_filePath, msg);
        }
    }

    /**
     * 菲律宾发票数据处理
     *
     * @param companyCode
     * @param filePath
     * @param OCR_filePath
     * @param ETL_filePath
     * @param fileName
     * @throws Exception
     */
    public static void clearPHIFunction(String companyCode, String filePath, String OCR_filePath, String ETL_filePath, String fileName) throws Exception {
        String invoiceReferenceNumber = "";
        String totalAmount = "";
        String currency = "";
        String invoiceDate = "";
        String posoNumber = "";
        File file = new File(filePath);
        byte[] fileData = FileUtils.readFileToByteArray(file);//读取图片
        //OCR模板扫描 返回结果数据
        String json = ocrImageFile("APSigPhi", file.getName(), fileData);
        System.out.println("输出菲律宾json:" + json);
        JSONObject jsonObject = JSON.parseObject(json);
        if (jsonObject.get("result").toString().contains("success")) {
            //处理json数据
            JSONArray jsonArray = jsonObject.getJSONObject("ocrResult").getJSONArray("ranges");
            Map<String, Object> map = new HashMap<>();
            for (int i = 0; i < jsonArray.size(); i++) {
                JSONObject object = (JSONObject) jsonArray.get(i);
                switch (object.get("rangeId").toString()) {
                    case "InvoiceDate":
                        if (object.get("value") != null) {
                            invoiceDate = clearInvoiceDate(object.get("value").toString());
                        }
                        break;
                    case "PurchaseOrderNumber":
                        if (object.get("value") != null) {
                            posoNumber = object.get("value").toString();
                        }
                        break;
                    case "InvoiceReferenceNumber":
                        if (object.get("value") != null) {
                            invoiceReferenceNumber = object.get("value").toString();
                        }
                        break;
                    case "POShortText":
                        if (object.get("value") != null) {
                            String[] val = object.get("value").toString().split(" ");
                            List<String> list = Arrays.asList(val);
                            map.put("POShortText", list);
                        }
                        break;
                    case "Quantity":
                        if (object.get("value") != null) {
                            String[] vl = object.get("value").toString().split(" ");
                            List<String> lists = Arrays.asList(vl);
                            List<String> list1 = new ArrayList<>();
                            for (String str : lists) {
                                if (str.matches("[0-9]{1,}")) {
                                    list1.add(str);
                                }
                            }
                            map.put("Quantity", list1);
                        } else {
                            List<String> list1 = new ArrayList<>();
                            map.put("Quantity", list1);
                        }
                        break;
                    case "Price":
                        if (object.get("value") != null) {
                            String[] vl = object.get("value").toString().split(" ");
                            List<String> lists = Arrays.asList(vl);
                            List<String> list1 = new ArrayList<>();
                            for (String str : lists) {
                                Pattern p = Pattern.compile(".*\\d+.*");
                                Matcher m = p.matcher(str);
                                if (m.matches()) {
                                    list1.add(str);
                                }
                            }
                            map.put("UnitPrice", list1);
                        } else {
                            List<String> list1 = new ArrayList<>();
                            map.put("UnitPrice", list1);
                        }
                        break;
                    case "Amount":
                        if (object.get("value") != null) {
                            String value = object.get("value").toString().replace(" ", "");
                            List<String> list1 = new ArrayList<>();
                            String[] regerSplit = REGER_CURRENCY;//获取货币号数组
                            for (String curr : regerSplit) {
                                if (value.contains(curr)) {
                                    String[] vl = value.split(curr);
                                    List<String> list = Arrays.asList(vl);
                                    for (int j = 1; j < list.size(); j++) {
                                        String amount = clearData(list.get(j));
                                        if (!"".equals(amount)) {
                                            list1.add(amount);
                                        }
                                    }
                                    currency = curr;
                                    break;
                                }
                            }
                            map.put("Amount", list1);
                        } else {
                            List<String> list = new ArrayList<>();
                            map.put("Amount", list);
                        }
                        break;
                    case "TotalAmount":
                        if (object.get("value") != null) {
                            String value = object.get("value").toString().replace(" ", "");
                            totalAmount = object.get("value").toString();
                            String[] regerSplit = REGER_CURRENCY;//获取货币号数组
                            for (String curr : regerSplit) {
                                if (value.contains(curr)) {
                                    String[] vl = value.split(curr);
                                    List<String> list = Arrays.asList(vl);
                                    for (int j = 1; j < list.size(); j++) {
                                        String str = clearData(list.get(j));
                                        if (!"".equals(str)) {
                                            totalAmount = str;
                                        }
                                    }
                                }
                            }
                        }
                        break;

                }
            }
            System.out.println("invoiceReferenceNumber:" + invoiceReferenceNumber);
            System.out.println("currency:" + currency);
            System.out.println("invoiceDate:" + invoiceDate);
            System.out.println("posoNumber:" + posoNumber);
            System.out.println("totalamount:" + totalAmount);
            for (String key : map.keySet()) {//查看
                System.out.println(key + ":" + map.get(key));
            }
            List<String> amountList = (List<String>) map.get("Amount");
            List<String> priceList = (List<String>) map.get("UnitPrice");
            List<String> quantityList = (List<String>) map.get("Quantity");
            List<String> POShortTextList = (List<String>) map.get("POShortText");
            if (quantityList.size() > 0 && quantityList.size() == priceList.size() && quantityList.size() == amountList.size() && quantityList.size() == POShortTextList.size()) {
                List<DataDO> dataDOList = new ArrayList<>();
                for (int i = 0; i < quantityList.size(); i++) {
                    DataDO dataDO = new DataDO();
                    dataDO.setDownloadStatus("OK");
                    dataDO.setCompanyCode(companyCode);
                    dataDO.setFilepath(fileName);
                    dataDO.setInvoicedate(invoiceDate);
                    dataDO.setAmount(convertData(amountList.get(i)));
                    dataDO.setInvoiceReferenceNumber(invoiceReferenceNumber);
                    dataDO.setPurchaseOrderNumber(posoNumber);
                    dataDO.setPoshorttext(POShortTextList.get(i));
                    dataDO.setQuantity(quantityList.get(i));
                    dataDO.setTotalAmount(convertData(totalAmount));
                    dataDO.setUnitPrice(priceList.get(i));
                    dataDO.setOCRStatus("OK");
                    dataDO.setCurrency(currency);
                    dataDOList.add(dataDO);
                }
                //生成ETL表
                excelOutput_OCR(dataDOList, ETL_filePath);
                //生成OCR表
                excelOutput_OCR(dataDOList, OCR_filePath);
            } else {
                String msg = "OCR Data Incompleted";
                templateNot_SGP(companyCode, fileName, ETL_filePath, OCR_filePath, msg);
            }
        } else {
            String msg = "OCR Template not match";
            templateNot_SGP(companyCode, fileName, ETL_filePath, OCR_filePath, msg);
        }
    }

    /**
     * 法国or比利时发票数据处理
     *
     * @param companyCode
     * @param filePath
     * @param OCR_filePath
     * @param ETL_filePath
     * @param fileName
     * @throws Exception
     */
    public static void clearFAR_BELFunction(String companyCode, String filePath, String OCR_filePath, String ETL_filePath, String fileName) throws Exception {
        File file = new File(filePath);
        byte[] fileData = FileUtils.readFileToByteArray(file);//读取图片
        //使用第一个OCR模板扫描 返回结果数据
        String json = ocrImageFile("APSigUnified_0109", file.getName(), fileData);
        JSONObject jsonObject = JSON.parseObject(json);
        //判断结果数据返回情况
        if (jsonObject.get("result").toString().contains("success")) {
            System.out.println("输出法国或者比利时json:" + json);
            clearAPSigUnifiedOne(jsonObject, companyCode, filePath, OCR_filePath, ETL_filePath, fileName);
        } else {
            //若模板以返回结果不成功，则使用第二个模板进行扫描
            String json2 = ocrImageFile("APSigUnified_Template2", file.getName(), fileData);
            System.out.println("输出法国或者比利时json2:" + json2);
            JSONObject jsonObject2 = JSON.parseObject(json2);
            if (jsonObject2.get("result").toString().contains("success")) {
                clearAPSigUnifiedTwo(jsonObject2, companyCode, filePath, OCR_filePath, ETL_filePath, fileName);
            } else {
                String msg = "OCR Template not match";
                templateNot_SGP(companyCode, fileName, ETL_filePath, OCR_filePath, msg);
            }
        }
    }

    /**
     * OCR 模板扫描返回json数据处理
     *
     * @param jsonObject
     * @param companyCode
     * @param filePath
     * @param OCR_filePath
     * @param ETL_filePath
     * @param fileName
     */
    public static void clearAPSigUnifiedOne(JSONObject jsonObject, String companyCode, String filePath, String OCR_filePath, String ETL_filePath, String fileName) {
        String invoiceReferenceNumber = "";
        String totalAmount = "";
        String currency = "";
        String invoiceDate = "";
        String posoNumber = "";
        JSONArray jsonArray = jsonObject.getJSONObject("ocrResult").getJSONArray("ranges");
        Map<String, Object> map = new HashMap<>();
        for (int i = 0; i < jsonArray.size(); i++) {
            JSONObject object = (JSONObject) jsonArray.get(i);
            switch (object.get("rangeId").toString()) {
                case "InvoiceDate":
                    if (object.get("value") != null) {
                        invoiceDate = clearInvoiceDate(object.get("value").toString());
                    }
                    break;
                case "InvoiceReferenceNumber":
                    if (object.get("value") != null) {
                        invoiceReferenceNumber = clearIRNumber(object.get("value").toString());
                    }
                    break;
                case "POShortText":
                    if (object.get("value") != null) {
                        String[] vl = object.get("value").toString().split(" ");
                        List<String> list = Arrays.asList(vl);
                        List arrList = new ArrayList(list);
                        List<String> arrLists = new ArrayList<>();
                        for (int j = 0; j < arrList.size(); j++) {
                            if ("CHINA".equals(arrList.get(j))) {
                                arrList.remove(arrList.get(j));
                                arrList.remove(arrList.get(j - 1));
                                arrList.remove(arrList.get(j - 2));
                            }
                        }
                        for (int j = 0; j < arrList.size(); j++) {
                            if (arrList.get(j).toString().length() >= 5) {
                                String str = checkPOShortText(arrList.get(j).toString());
                                arrLists.add(str);
                            }
                        }
                        map.put("POShortText", arrLists);
                    }
                    break;
                case "POSONumber":
                    if (object.get("value") != null) {
                        posoNumber = checkPOSO(object.get("value").toString());
                        posoNumber = checkPONumber(posoNumber);
                    }
                    break;
                case "Quantity":
                    if (object.get("value") != null) {
                        String[] vl = object.get("value").toString().split(" ");
                        List<String> list = Arrays.asList(vl);
                        List<String> list1 = new ArrayList<>();
                        for (String quantity : list) {
                            quantity = checkData(quantity);
                            if (quantity.contains(QUANTITY_SGP)) {
                                Pattern p = Pattern.compile(REGER);//提取有效数字
                                quantity = p.matcher(quantity).replaceAll("").trim();
                                if (!"".equals(quantity)) {
                                    list1.add(quantity);
                                }
                            }
                        }
                        map.put("Quantity", list1);
                    } else {
                        List<String> list = new ArrayList<>();
                        map.put("Quantity", list);
                    }
                    break;
                case "Price":
                    if (object.get("value") != null) {
                        String unitPrice = object.get("value").toString().replace(" ", "").replace("CBP", "GBP");
                        List<String> list1 = new ArrayList<>();
                        String[] regerSplit = REGER_CURRENCY;//获取货币号数组
                        for (String curr : regerSplit) {
                            if (unitPrice.contains(curr)) {
                                String[] vl = unitPrice.split(curr);
                                List<String> list = Arrays.asList(vl);
                                for (int j = 1; j < list.size(); j++) {
                                    String price = clearData(list.get(j));
                                    if (!"".equals(price)) {
                                        list1.add(price);
                                    }
                                }
                                currency = curr;
                                break;
                            }
                        }
                        map.put("UnitPrice", list1);
                    } else {
                        List<String> list = new ArrayList<>();
                        map.put("UnitPrice", list);
                    }
                    break;
                //模板扫描时，amount以及totalAmount同时获得
                case "Amount":
                    if (object.get("value") != null) {
                        String[] vl = object.get("value").toString().split(" ");
                        List<String> list = Arrays.asList(vl);
                        List<Double> list1 = new ArrayList<>();
                        for (String amount : list) {
                            amount = checkData(amount);
                            if (amount.matches(REGER_SGP)) {
                                amount = amount.replaceAll("[a-zA-Z]", "").trim();
                                list1.add(Double.parseDouble(amount));
                            }
                        }
                        map.put("Amount", list1);
                    } else {
                        List<String> list = new ArrayList<>();
                        map.put("Amount", list);
                    }
                    break;
            }
        }
        for (String key : map.keySet()) {//查看
            System.out.println(key + ":" + map.get(key));
        }
        System.out.println("输出currency：" + currency);
        //获取集合，判断集合中的最大值，即为totalAmount，并将最大值从集合删除
        List<Double> amountList = (List<Double>) map.get("Amount");
        if (amountList.size() > 0) {
            totalAmount = String.valueOf(Collections.max(amountList));
            System.out.println("输出最大值：" + totalAmount);
            int j = amountList.lastIndexOf(Double.parseDouble(totalAmount));
            amountList.remove(j);
        }
        List<String> priceList = (List<String>) map.get("UnitPrice");
        List<String> quantityList = (List<String>) map.get("Quantity");
        List<String> POShortTextList = (List<String>) map.get("POShortText");
        //判断提取获得的数据数量是否相等，不相等即为模板不适用
        if (quantityList.size() > 0 && quantityList.size() == priceList.size() && quantityList.size() == amountList.size() && quantityList.size() == POShortTextList.size()) {
            List<DataDO> dataDOList = new ArrayList<>();
            for (int i = 0; i < quantityList.size(); i++) {
                DataDO dataDO = new DataDO();
                dataDO.setDownloadStatus("OK");
                dataDO.setCompanyCode(companyCode);
                dataDO.setFilepath(fileName);
                dataDO.setInvoicedate(invoiceDate);
                dataDO.setAmount(convertData(String.valueOf(amountList.get(i))));
                dataDO.setInvoiceReferenceNumber(invoiceReferenceNumber);
                dataDO.setPurchaseOrderNumber(posoNumber);
                dataDO.setPoshorttext(POShortTextList.get(i));
                dataDO.setQuantity(quantityList.get(i));
                dataDO.setTotalAmount(convertData(totalAmount));
                dataDO.setUnitPrice(priceList.get(i));
                dataDO.setOCRStatus("OK");
                dataDO.setCurrency(currency);
                dataDOList.add(dataDO);
            }
            //生成ETL表
            excelOutput_OCR(dataDOList, ETL_filePath);
            //生成OCR表
            excelOutput_OCR(dataDOList, OCR_filePath);
        } else {
            String msg = "OCR Data Incompleted";
            templateNot_SGP(companyCode, fileName, ETL_filePath, OCR_filePath, msg);
        }
    }

    public static void clearAPSigUnifiedTwo(JSONObject jsonObject, String companyCode, String filePath, String OCR_filePath, String ETL_filePath, String fileName) {
        String invoiceDate = "";
        String purchaseOrderNumber = "";
        String invoiceReferenceNumber = "";
        String currency = "";
        String totalAmount = "";
        System.out.println("输出模板二数据：" + jsonObject);
        JSONArray jsonArray = jsonObject.getJSONObject("ocrResult").getJSONArray("ranges");
        Map<String, Object> map = new HashMap<>();
        for (int i = 0; i < jsonArray.size(); i++) {
            JSONObject object = (JSONObject) jsonArray.get(i);
            switch (object.get("rangeId").toString()) {
                case "InvoiceDate":
                    if (object.get("value") != null) {
                        invoiceDate = clearInvoiceDate(object.get("value").toString());
                    }
                    break;
                case "PurchaseOrderNumber":
                    if (object.get("value") != null) {
                        purchaseOrderNumber = checkPOSO(object.get("value").toString());
                        purchaseOrderNumber = checkPONumber(purchaseOrderNumber);
                    }
                    break;
                case "InvoiceReferenceNumber":
                    if (object.get("value") != null) {
                        invoiceReferenceNumber = object.get("value").toString();
                        invoiceReferenceNumber = invoiceReferenceNumber.replaceAll("[\u4e00-\u9fa5]", "").trim();
                    }
                    break;
                case "POShortText":
                    if (object.get("value") != null) {
                        String[] vl = object.get("value").toString().split(" ");
                        List<String> list = Arrays.asList(vl);
                        map.put("POShortText", list);
                    }
                    break;
                case "Quantity":
                    if (object.get("value") != null) {
                        String[] vl = object.get("value").toString().split(" ");
                        List<String> list = Arrays.asList(vl);
                        List<String> list1 = new ArrayList<>();
                        for (String str : list) {
                            if (str.matches("^([0-9]+[a-zA-Z]{3})+$")) {
                                str = str.replaceAll("[a-zA-Z]", "").trim();
                                list1.add(str);
                            }
                        }
                        map.put("Quantity", list1);
                    } else {
                        List<String> list = new ArrayList<>();
                        map.put("Quantity", list);
                    }
                    break;
                case "Price":
                    if (object.get("value") != null) {
                        String unitPrice = object.get("value").toString().replace(" ", "").replace("CBP", "GBP");
                        List<String> list1 = new ArrayList<>();
                        String[] regerSplit = REGER_CURRENCY;//获取货币号数组
                        for (String curr : regerSplit) {
                            if (unitPrice.contains(curr)) {
                                String[] vl = unitPrice.split(curr);
                                List<String> list = Arrays.asList(vl);
                                for (int j = 1; j < list.size(); j++) {
                                    String price = clearData(list.get(j));
                                    if (!"".equals(price)) {
                                        list1.add(price);
                                    }
                                }
                                currency = curr;
                                break;
                            }
                        }
                        map.put("UnitPrice", list1);
                    } else {
                        List<String> list = new ArrayList<>();
                        map.put("UnitPrice", list);
                    }
                    break;
                case "Amount":
                    if (object.get("value") != null) {
                        String value = object.get("value").toString().replace(" ", "");
                        List<Double> list1 = new ArrayList<>();
                        String[] regerSplit = REGER_CURRENCY;//获取货币号数组
                        for (String curr : regerSplit) {
                            if (value.contains(curr)) {
                                String[] vl = value.split(curr);
                                List<String> list = Arrays.asList(vl);
                                for (int j = 1; j < list.size(); j++) {
                                    String amount = clearData(list.get(j));
                                    if (!"".equals(amount)) {
                                        list1.add(Double.parseDouble(amount));
                                    }
                                }
                                break;
                            }
                        }
                        map.put("Amount", list1);
                    } else {
                        List<String> list = new ArrayList<>();
                        map.put("Amount", list);
                    }
                    break;
            }
        }
        for (String key : map.keySet()) {//查看
            System.out.println(key + ":" + map.get(key));
        }
        System.out.println("输出currency：" + currency);
        System.out.println("invoiceReferenceNumber:" + invoiceReferenceNumber);
        System.out.println("invoiceDate:" + invoiceDate);
        System.out.println("posoNumber:" + purchaseOrderNumber);
        //获取集合，判断集合中的最大值，即为totalAmount，并将最大值从集合删除
        List<Double> amountList = (List<Double>) map.get("Amount");
        if (amountList.size() > 0) {
            totalAmount = String.valueOf(Collections.max(amountList));
            System.out.println("输出最大值：" + totalAmount);
            int j = amountList.lastIndexOf(Double.parseDouble(totalAmount));
            amountList.remove(j);
        }
        List<String> priceList = (List<String>) map.get("UnitPrice");
        List<String> quantityList = (List<String>) map.get("Quantity");
        List<String> POShortTextList = (List<String>) map.get("POShortText");
        //判断提取获得的数据数量是否相等，不相等即为模板不适用
        if (quantityList.size() > 0 && quantityList.size() == priceList.size() && quantityList.size() == amountList.size() && quantityList.size() == POShortTextList.size()) {
            List<DataDO> dataDOList = new ArrayList<>();
            for (int i = 0; i < quantityList.size(); i++) {
                DataDO dataDO = new DataDO();
                dataDO.setDownloadStatus("OK");
                dataDO.setCompanyCode(companyCode);
                dataDO.setFilepath(fileName);
                dataDO.setInvoicedate(invoiceDate);
                dataDO.setAmount(convertData(String.valueOf(amountList.get(i))));
                dataDO.setInvoiceReferenceNumber(invoiceReferenceNumber);
                dataDO.setPurchaseOrderNumber(purchaseOrderNumber);
                dataDO.setPoshorttext(POShortTextList.get(i));
                dataDO.setQuantity(quantityList.get(i));
                dataDO.setTotalAmount(convertData(totalAmount));
                dataDO.setUnitPrice(priceList.get(i));
                dataDO.setOCRStatus("OK");
                dataDO.setCurrency(currency);
                dataDOList.add(dataDO);
            }
            //生成ETL表
            excelOutput_OCR(dataDOList, ETL_filePath);
            //生成OCR表
            excelOutput_OCR(dataDOList, OCR_filePath);
        } else {
            String msg = "OCR Data Incompleted";
            templateNot_SGP(companyCode, fileName, ETL_filePath, OCR_filePath, msg);
        }
    }
}
