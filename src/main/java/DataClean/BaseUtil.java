package DataClean;

import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.ContentType;
import org.apache.http.entity.mime.HttpMultipartMode;
import org.apache.http.entity.mime.MultipartEntityBuilder;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.util.EntityUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.math.BigDecimal;
import java.nio.charset.Charset;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Pattern;

/**
 * 数据处理公共类
 */
public class BaseUtil {
    public static final String REGER = "[^，,.>0-9]";
    public static final String[] REGER_CURRENCY = {"USD","SD","EUR","GBP","JPY","CNY","HKD","AED","AUD"};//货币号
    /**
     * 该方法用于OCR模板扫面发票，获取所需数据
     * @param imageType
     * @param fileName
     * @param fileData
     * @return
     * @throws Exception
     */
    public static String ocrImageFile(String imageType,String fileName,byte[] fileData) throws Exception{
        String baseUrl = "http://ocrserver.openserver.cn:8090/OcrServer/ocr/ocrImageByTemplate";
//        String baseUrl = "http://10.138.93.103:8080/OcrServer/ocr/ocrImageByTemplate";
        HttpPost post = new HttpPost(baseUrl);
        ContentType contentType = ContentType.create("multipart/form-data", Charset.forName("UTF-8"));
        MultipartEntityBuilder builder = MultipartEntityBuilder.create();
        builder.setMode(HttpMultipartMode.BROWSER_COMPATIBLE);
        builder.setCharset(Charset.forName("UTF-8"));
        builder.addBinaryBody("file", fileData, contentType, fileName);// 文件流
        builder.addTextBody("imageType", imageType,contentType);// 类似浏览器表单提交，对应input的name和value
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

    /**
     * 清洗单号
     * @param IRNumber
     * @return
     */
    public static String clearIRNumber(String IRNumber){
        if (!"".equals(IRNumber)){
            IRNumber = IRNumber.replace("CQ","C0");
            IRNumber = IRNumber.replace("HQ","H0");
            IRNumber = IRNumber.replace("HQQ","H00");
            IRNumber = IRNumber.replace("H0Q","H00");
            IRNumber = IRNumber.trim();
            if (IRNumber.contains(" ")){
                IRNumber = IRNumber.substring(0,IRNumber.indexOf(" "));
            }
            StringBuilder sb = new StringBuilder(IRNumber);
//            sb.replace(0, 2, "C0");
            String str = sb.substring(6,8);
            switch (str){
                case "BO":
                    sb.replace(6, 8, "B0");
                    break;
                case "CB":
                    sb.replace(6, 8, "GB");
                    break;
                case "1T":
                    sb.replace(6, 8, "IT");
                    break;
                case "1P":
                    sb.replace(6, 8, "JP");
                    break;
                case "IP":
                    sb.replace(6, 8, "JP");
                    break;
                case "jP":
                    sb.replace(6, 8, "JP");
                    break;
                case "jp":
                    sb.replace(6, 8, "JP");
                    break;
                case "」P":
                    sb.replace(6, 8, "JP");
                    break;
            }
            if (IRNumber.length() >= 12){
                if ("SH".equals(sb.substring(10,12))){
                    sb.replace(10, 12, "5H");
                }
            }
            String str1 = sb.toString().substring(8);
            str1 = str1.replace("O","0");
            str1 = str1.replace("Q","0");
            IRNumber = sb.toString().substring(0,8) + str1;
        }
        return IRNumber;
    }

    /**
     * 数据处理(必须使用货币号拆分出每条数据时使用)
     * @param data
     * @return
     */
    public static String clearData(String data){
        if (!"".equals(data)){
            data = data.replace("U","11");
            data = data.replace("O","0");
            data = data.replace("*",".");
            Pattern p = Pattern.compile(REGER);//提取有效数字
            data = p.matcher(data).replaceAll("").trim();
            if (!"".equals(data)){//获取到的数据不能为空
                data = data.replaceAll("(?:,|>)",".");//将标点换成小数点
                if (data.contains(".")){//判断小数点是否存在
                    String str = data.substring(data.indexOf(".",0) + 1);//截取小数点之后的字符串
                    if (str.length() >= 2){
                        data = data.substring(0, data.indexOf(".",0) + 3);
                    }
                }
            }
        }
        return data;
    }

    /**
     * 各国发票数据生成OCR表
     * @param dataDOS
     * @param fileName
     */
    public static void excelOutput_OCR(List<DataDO> dataDOS, String fileName) {
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        XSSFSheet xssfSheet = xssfWorkbook.createSheet("Sheet1");
        Row row0 = xssfSheet.createRow(0);
        String[] headers = new String[]{"FileName","DownloadStatus","CompanyCode", "Amount", "InvoiceDate", "InvoiceReferenceNumber","InvoiceReferenceNumber2", "POShortText", "PurchaseOrderNumber", "Quantity", "TaxAmount", "TotalAmount", "UnitPrice","Currency","SONumber","GoodsDescription","OCRStatus"};
        for (int i = 0; i < headers.length; i++) {
            xssfSheet.setColumnWidth(i, 5000);
            XSSFCell cell = (XSSFCell) row0.createCell(i);
            XSSFRichTextString text = new XSSFRichTextString(headers[i]);
            cell.setCellValue(text);
        }
        int rowNum = 1;
        for (DataDO dataDO : dataDOS){
            XSSFRow row = xssfSheet.createRow(rowNum);
            row.createCell(0).setCellValue(dataDO.getFilepath());
            row.createCell(1).setCellValue(dataDO.getDownloadStatus());
            row.createCell(2).setCellValue(dataDO.getCompanyCode());
            row.createCell(3).setCellValue(dataDO.getAmount());
            row.createCell(4).setCellValue(dataDO.getInvoicedate());
            row.createCell(5).setCellValue(dataDO.getInvoiceReferenceNumber());
            row.createCell(6).setCellValue(dataDO.getInvoiceReferenceNumber2());
            row.createCell(7).setCellValue(dataDO.getPoshorttext());
            row.createCell(8).setCellValue(dataDO.getPurchaseOrderNumber());
            row.createCell(9).setCellValue(dataDO.getQuantity());
            row.createCell(10).setCellValue(dataDO.getTaxAmount());
            row.createCell(11).setCellValue(dataDO.getTotalAmount());
            row.createCell(12).setCellValue(dataDO.getUnitPrice());
            row.createCell(13).setCellValue(dataDO.getCurrency());
            row.createCell(14).setCellValue(dataDO.getSOnumber());
            row.createCell(15).setCellValue(dataDO.getGoodDescription());
            row.createCell(16).setCellValue(dataDO.getOCRStatus());
            rowNum++;
        }
        try {
            FileOutputStream fileOutputStream=new FileOutputStream(fileName);
            xssfWorkbook.write(fileOutputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 模板不适用
     * @param companyCode
     * @param fileName
     * @param OCR_filePath
     */
    public static void templateNot(String companyCode, String fileName, String OCR_filePath,String message){
        List<DataDO> dataDOList = new ArrayList<>();
        DataDO dataDO = new DataDO();
        dataDO.setOCRStatus(message);
        dataDO.setCompanyCode(companyCode);
        dataDO.setDownloadStatus("OK");
        dataDO.setFilepath(fileName);
        dataDOList.add(dataDO);
        excelOutput_OCR(dataDOList, OCR_filePath);
    }

    /**
     * 保留两位有效数字
     * @param data
     * @return
     */
    public static String convertData(String data){
        if (!"".equals(data) && data != null ){
            BigDecimal bd = new BigDecimal(data);
            data = bd.setScale(2, BigDecimal.ROUND_HALF_UP).toPlainString();
        }
        return data;
    }

    /**
     * AR核销数据处理生成SAP表
     * @param middleEastSAPList
     */
    public static void excelOutput_SAP_AR( List<MiddleEastSAP> middleEastSAPList, String filePath) {
        System.out.println(filePath);
        //创建表格
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        //定义第一个sheet页
        XSSFSheet xssfSheet = xssfWorkbook.createSheet("Sheet1");
        //第一个sheet页数据（生成台账表）
        Row row0 = xssfSheet.createRow(0);
        String[] headers = new String[]{"ID","No", "Document date", "Type", "Company Code", "Posting Date", "Period", "Currency/Rate", "Document Number", "Translatn Date", "Reference", "Cross-CC No.", "Doc.Header Text", "Branch Number", "Trading Part.BA", "Postkey", "GL", "SGL Ind", "Amount", "Tax Code", "Business Area", "Cost Center", "Profit Center", "Value Date", "Due On","Bline Date","Issue Date","Ext. No","Bank/Acct No","Disc Base","Assignment","Text","Long Text","Reason code"};
        for (int i = 0; i < headers.length; i++) {
            xssfSheet.setColumnWidth(i+1, 5000);
            XSSFCell cell = (XSSFCell) row0.createCell(i);
            XSSFRichTextString text = new XSSFRichTextString(headers[i]);
            cell.setCellValue(text);
        }
        int rowNum = 1;
        for (MiddleEastSAP MEastSap : middleEastSAPList){
            XSSFRow row = xssfSheet.createRow(rowNum);
            xssfSheet.setColumnWidth(rowNum, 5000);
            row.createCell(0).setCellValue(MEastSap.getId());
            row.createCell(1).setCellValue(MEastSap.getNo());
            row.createCell(2).setCellValue(MEastSap.getDocumentDate());
            row.createCell(3).setCellValue(MEastSap.getType());
            row.createCell(4).setCellValue(MEastSap.getCompanyCode());
            row.createCell(5).setCellValue(MEastSap.getPostingDate());
            row.createCell(6).setCellValue(MEastSap.getPeriod());
            row.createCell(7).setCellValue(MEastSap.getCurrencyRate());
            row.createCell(8).setCellValue(MEastSap.getDocumentNumber());
            row.createCell(9).setCellValue(MEastSap.getTranslatnDate());
            row.createCell(10).setCellValue(MEastSap.getReference());
            row.createCell(11).setCellValue(MEastSap.getCrossCCNo());
            row.createCell(12).setCellValue(MEastSap.getDocHeaderText());
            row.createCell(13).setCellValue(MEastSap.getBranchNumber());
            row.createCell(14).setCellValue(MEastSap.getTradingPartBA());
            row.createCell(15).setCellValue(MEastSap.getPostKey());
            row.createCell(16).setCellValue(MEastSap.getGl());
            row.createCell(17).setCellValue(MEastSap.getSglInd());
            String amount = MEastSap.getAmount();
            if (!"".equals(amount) && amount != null ){
                Pattern p = Pattern.compile("[^.0-9]");//提取有效数字
                amount = p.matcher(amount).replaceAll("").trim();
                row.createCell(18).setCellValue(convertData(amount));
            }
            row.createCell(19).setCellValue(MEastSap.getTaxCode());
            row.createCell(20).setCellValue(MEastSap.getBusinessArea());
            row.createCell(21).setCellValue(MEastSap.getCostCenter());
            row.createCell(22).setCellValue(MEastSap.getProfitCenter());
            row.createCell(23).setCellValue(MEastSap.getValueDate());
            row.createCell(24).setCellValue(MEastSap.getDuOn());
            row.createCell(25).setCellValue(MEastSap.getBlinDate());
            row.createCell(26).setCellValue(MEastSap.getIssueDate());
            row.createCell(27).setCellValue(MEastSap.getExtNo());
            row.createCell(28).setCellValue(MEastSap.getBankAcctNo());
            row.createCell(29).setCellValue(MEastSap.getDiscBase());
            row.createCell(30).setCellValue(MEastSap.getAssignment());
            row.createCell(31).setCellValue(MEastSap.getText());
            row.createCell(32).setCellValue(MEastSap.getLongText());
            row.createCell(33).setCellValue(MEastSap.getReasonCode());
            rowNum ++;
        }
        try {
            FileOutputStream fileOutputStream = new FileOutputStream(filePath);
            xssfWorkbook.write(fileOutputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 获取CostCentre数据
     * @param filePath_Cost
     * @return
     * @throws Exception
     */
    public static List getCostCenterExcel( String filePath_Cost) throws Exception{
        String productCode = "";
        String costCenter = "";
        File excelFile_cost = new File(filePath_Cost);
        FileInputStream EIL_file_IO = new FileInputStream(excelFile_cost);
        Workbook wb =  WorkbookFactory.create(EIL_file_IO);
        Sheet sheet = wb.getSheet("Product Cost Center");
        int coloumNum = sheet.getRow(0).getPhysicalNumberOfCells();
        //获取表头行
        Row row = sheet.getRow(0);
        List<CostCenter> costCenterList = new ArrayList<>();
        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row rows = sheet.getRow(r);//获取第r行
            CostCenter cost = new CostCenter();
            for (int i = 0; i < coloumNum ; i ++){
                row.getCell(i).setCellType(Cell.CELL_TYPE_STRING);
                if (rows.getCell(i) != null){
                    rows.getCell(i).setCellType(Cell.CELL_TYPE_STRING);
                    //获取表头行某个字段列，对应当前行的列数据，
                    if ("Product Code".equals(row.getCell(i).getStringCellValue())){
                        productCode = rows.getCell(i).getStringCellValue();
                        cost.setProductCode(productCode);
                    }
                    if ("Cost Center".equals(row.getCell(i).getStringCellValue())){
                        costCenter = rows.getCell(i).getStringCellValue();
                        cost.setCostCenter(costCenter);
                    }
                }
            }
            costCenterList.add(cost);
        }
        return costCenterList;
    }

    /**
     * 获取产品线productCode，所对应的costCenter
     * @param productCode
     * @param excelFile_cost
     * @return
     * @throws Exception
     */
    public static String getCostCenter(String productCode, String excelFile_cost) throws Exception{
        String costCenter = "";
        if (!"".equals(productCode) && productCode != null){
            //获取CostCenter 表数据
            List<CostCenter> costCenterList = getCostCenterExcel(excelFile_cost);
            for (CostCenter cost: costCenterList){
                if (productCode.equals(cost.getProductCode())){
                    costCenter = cost.getCostCenter();
                }
            }
        }
        return costCenter;
    }

    /**
     * 获取Account数据
     * @param filePath_Cost
     * @return
     * @throws Exception
     */
    public static List getAccountExcel( String filePath_Cost) throws Exception{
        String bank = "";
        String bankCurrency = "";
        String currency = "";
        String account = "";
        File excelFile_cost = new File(filePath_Cost);
        FileInputStream EIL_file_IO = new FileInputStream(excelFile_cost);
        Workbook wb =  WorkbookFactory.create(EIL_file_IO);
        Sheet sheet = wb.getSheet("Bank Account");
        int coloumNum = sheet.getRow(0).getPhysicalNumberOfCells();
        //获取表头行
        Row row = sheet.getRow(0);
        List<Account> accountList = new ArrayList<>();
        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row rows = sheet.getRow(r);//获取第r行
            Account account1 = new Account();
            for (int i = 0; i < coloumNum ; i ++){
                row.getCell(i).setCellType(Cell.CELL_TYPE_STRING);
                if (rows.getCell(i) != null){
                    rows.getCell(i).setCellType(Cell.CELL_TYPE_STRING);
                    //获取表头行某个字段列，对应当前行的列数据，
                    if ("Bank".equals(row.getCell(i).getStringCellValue())){
                        bank = rows.getCell(i).getStringCellValue();
                        account1.setBank(bank);
                    }
                    if ("Bank-Currency".equals(row.getCell(i).getStringCellValue())){
                        bankCurrency = rows.getCell(i).getStringCellValue();
                        account1.setBankCurrency(bankCurrency);
                    }
                    if ("Currency".equals(row.getCell(i).getStringCellValue())){
                        currency = rows.getCell(i).getStringCellValue();
                        account1.setCurrency(currency);
                    }
                    if ("Account".equals(row.getCell(i).getStringCellValue())){
                        account = rows.getCell(i).getStringCellValue();
                        account1.setAccount(account);
                    }
                }
            }
            accountList.add(account1);
        }
        return accountList;
    }

    /**
     * 获取transferTo，account
     * @param transferTo
     * @param excelFile_cost
     * @return
     * @throws Exception
     */
    public static String getAccountDetails(String transferTo, String excelFile_cost) throws Exception{
        String account = "";
        if (!"".equals(transferTo) && transferTo != null){
            //Account 表数据
            List<Account> accountList = getAccountExcel(excelFile_cost);
            for (Account account1: accountList){
                if (transferTo.equals(account1.getBankCurrency())){
                    account = account1.getAccount();
                }
            }
        }
        return account;
    }

    /**
     * 获取Account Number 对应的sapCode
     * @param account
     * @param banklistPath
     * @return
     * @throws Exception
     */
    public static String getAccount(String account, String banklistPath) throws Exception{
        String sapCode = "";
        if (!"".equals(account) && account != null){
            //banklist 表数据
            List<BankList> bankList = getBankListExcel(banklistPath);
            for (BankList bankList1: bankList){
                if (account.equals(bankList1.getAccountNumber())){
                    sapCode = bankList1.getSapCode();
                }
            }
        }
        return sapCode;
    }
    /**
     * 获取Account Number 对应的sapCode
     * @param taxCode
     * @param banklistPath
     * @return
     * @throws Exception
     */
    public static String getrefundAccount(String taxCode, String banklistPath) throws Exception{
        String CustomerCode = "";
        if (!"".equals(taxCode) && taxCode != null){
            //CusList 表数据
            List<BankList> bankList = getCusExcel(banklistPath);
            for (BankList bankList1: bankList){
//
                if (taxCode.equals(bankList1.getTaxCode().trim())){
                    CustomerCode = bankList1.getCustomerCode();
                    System.out.println(CustomerCode);
                }
            }
        }
        return CustomerCode;
    }
    /**
     * 获取101 other Account Number 对应的sapCode
     * @param
     * @param banklistPath
     * @return
     * @throws Exception
     */
    public static List<String> getOtherAccount(String title, String banklistPath) throws Exception{
        List<String> list = new ArrayList<>();
        String vendorCode = "";
        String ReasonCode = "";
        if (!"".equals(title) && title != null){
            //banklist 表数据
            List<BankList> bankList = getOtherBankListExcel(banklistPath);
            for (BankList bankList1: bankList){
//                System.out.println("account:"+title+",银行账号:"+bankList1.getVendorName());

                if (title.contains(bankList1.getVendorName())){
//                    System.out.println("匹配成功");
                    ReasonCode = bankList1.getSapCode();
                    vendorCode = bankList1.getAccountNumber();
                    list.add(vendorCode);
                    list.add(ReasonCode);

                }
            }
        }
        return list;
    }

    /**
     * 101获取BankList表数据
     * @param banklistPath
     * @return
     * @throws Exception
     */
    public static List getBankListExcel( String banklistPath) throws Exception{
        File excelFile_cost = new File(banklistPath);
        FileInputStream EIL_file_IO = new FileInputStream(excelFile_cost);
        Workbook wb =  WorkbookFactory.create(EIL_file_IO);
        Sheet sheet = wb.getSheetAt(0);
        int coloumNum = sheet.getRow(0).getPhysicalNumberOfCells();
        //获取表头行
        Row row = sheet.getRow(0);
        List<BankList> bankLists = new ArrayList<>();
        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row rows = sheet.getRow(r);//获取第r行
            BankList bankList = new BankList();
            String sapCode = "";
            String accountNumber = "";
            for (int i = 0; i < coloumNum ; i ++){
                row.getCell(i).setCellType(Cell.CELL_TYPE_STRING);
                if (rows.getCell(i) != null){
                    rows.getCell(i).setCellType(Cell.CELL_TYPE_STRING);
                    //获取表头行某个字段列，对应当前行的列数据，
                    if ("Account number".equals(row.getCell(i).getStringCellValue())){
                        accountNumber = rows.getCell(i).getStringCellValue();
                        bankList.setAccountNumber(accountNumber);
                    }
                    if ("SAP Code".equals(row.getCell(i).getStringCellValue())){
                        sapCode = rows.getCell(i).getStringCellValue();
                        bankList.setSapCode(sapCode);
                    }
                }
            }
            bankLists.add(bankList);
        }
        return bankLists;

    }
    /**
     * 101从cusList表里面获取customercode
     * @param banklistPath
     * @return
     * @throws Exception
     */
    public static List getCusExcel( String banklistPath) throws Exception{
        File excelFile_cost = new File(banklistPath);
        FileInputStream EIL_file_IO = new FileInputStream(excelFile_cost);
        Workbook wb =  WorkbookFactory.create(EIL_file_IO);
        Sheet sheet = wb.getSheetAt(0);
        int coloumNum = sheet.getRow(0).getPhysicalNumberOfCells();
        //获取表头行
        Row row = sheet.getRow(0);

        List<BankList> bankLists = new ArrayList<>();
        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row rows = sheet.getRow(r);//获取第r行
            BankList bankList = new BankList();
            String CustomerCode = "";
            String TaxCode = "";
            for (int i = 0; i < coloumNum ; i ++){
                row.getCell(i).setCellType(Cell.CELL_TYPE_STRING);
                if (rows.getCell(i) != null){
                    rows.getCell(i).setCellType(Cell.CELL_TYPE_STRING);
                    //获取表头行某个字段列，对应当前行的列数据，
                    if ("Customer".equals(row.getCell(i).getStringCellValue())){
                        CustomerCode = rows.getCell(i).getStringCellValue();
                        bankList.setCustomerCode(CustomerCode);
                    }
                    if ("Tax Number 1".equals(row.getCell(i).getStringCellValue())){

                        TaxCode = rows.getCell(i).getStringCellValue();
                        bankList.setTaxCode(TaxCode);
                    }
                }
            }
            bankLists.add(bankList);
        }
        return bankLists;

    }
    /**
     * 101获取BankList other表数据
     * @param banklistPath
     * @return
     * @throws Exception
     */
    public static List getOtherBankListExcel( String banklistPath) throws Exception{
        File excelFile_cost = new File(banklistPath);
        FileInputStream EIL_file_IO = new FileInputStream(excelFile_cost);
        Workbook wb =  WorkbookFactory.create(EIL_file_IO);
        Sheet sheet = wb.getSheetAt(0);
        int coloumNum = sheet.getRow(0).getPhysicalNumberOfCells();
        //获取表头行
        Row row = sheet.getRow(0);
        List<BankList> bankLists = new ArrayList<>();
        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row rows = sheet.getRow(r);//获取第r行
            BankList bankList = new BankList();
            String sapCode = "";
            String accountNumber = "";
            String Vendorname = "";
            for (int i = 0; i < coloumNum ; i ++){
                row.getCell(i).setCellType(Cell.CELL_TYPE_STRING);
                if (rows.getCell(i) != null){
                    rows.getCell(i).setCellType(Cell.CELL_TYPE_STRING);
                    //获取表头行某个字段列，对应当前行的列数据，
                    if ("VendorName".equals(row.getCell(i).getStringCellValue())){
                        Vendorname = rows.getCell(i).getStringCellValue();
                        bankList.setVendorName(Vendorname);
                    }
                    if ("VendorCode".equals(row.getCell(i).getStringCellValue())){
                        accountNumber = rows.getCell(i).getStringCellValue();
                        bankList.setAccountNumber(accountNumber);
                    }
                    if ("ReasonCode".equals(row.getCell(i).getStringCellValue())){
                        sapCode = rows.getCell(i).getStringCellValue();
                        bankList.setSapCode(sapCode);
                    }
                }
            }
            bankLists.add(bankList);
        }
        return bankLists;

    }


    /**
     * 中东地区AR核销数据处理业务
     * 获取台账数据
     * @param filePath
     * @return
     * @throws Exception
     */
    public static List getLedgerExcelData(String filePath) throws Exception {
        File excelFile = new File(filePath);
        FileInputStream EIL_file_IO = new FileInputStream(excelFile);
        Workbook wb = WorkbookFactory.create(EIL_file_IO);
        Sheet sheet = wb.getSheetAt(0);
        SimpleDateFormat sdf1 = new SimpleDateFormat("MM/dd/yyyy");
        List<MiddleEastAll> MiddleEastList = new ArrayList<>();
        System.out.println("原来表里面有"+sheet.getLastRowNum()+"行");
        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            String id = "";
            String bank = "";
            String documentDate = "";
            String income = "";//正向
            String charge = "";//负向
            String currency = "";
            String summary = "";
            String ttLCMark = "";
            String customerName = "";
            String customerCode = "";
            String recognizedAmount = "";
            String productCode = "";
            String pi = "";
            String prepayment = "";
            String transferTo = "";
            String depositPrincipal = "";
            String depositInterest = "";
            String staffName = "";
            String remark = "";
            String sapIncomeNo = "";
            String comment = "";
            String sapNoForOther = "";
            String sapClearingNo = "";
            String EmailAddress = "";
            Row rows = sheet.getRow(r);//获取第r行
            if (rows.getCell(1) != null){
                rows.getCell(1).setCellType(Cell.CELL_TYPE_STRING);
                id = rows.getCell(1).getStringCellValue();
            }
            System.out.println("第"+r+"行,"+"获取的id："+id);

            if (rows.getCell(2) != null){
                rows.getCell(2).setCellType(Cell.CELL_TYPE_STRING);
                bank = rows.getCell(2).getStringCellValue();
            }
            if (rows.getCell(3) != null){
                if (rows.getCell(3).getCellType() == 0){
                    if (DateUtil.isCellDateFormatted(rows.getCell(3))){//判断是否是时间类型
                        documentDate = sdf1.format(rows.getCell(3).getDateCellValue());
                    }
                }else {
                    rows.getCell(3).setCellType(Cell.CELL_TYPE_STRING);
                    documentDate = rows.getCell(3).getStringCellValue();
                }
            }
            if (rows.getCell(4) != null){
                rows.getCell(4).setCellType(Cell.CELL_TYPE_STRING);
                income = rows.getCell(4).getStringCellValue();
            }
            if (rows.getCell(5) != null){
                rows.getCell(5).setCellType(Cell.CELL_TYPE_STRING);
                charge = rows.getCell(5).getStringCellValue();
            }
            if (rows.getCell(6) != null){
                rows.getCell(6).setCellType(Cell.CELL_TYPE_STRING);
                currency = rows.getCell(6).getStringCellValue();
            }
            if (rows.getCell(7) != null){
                rows.getCell(7).setCellType(Cell.CELL_TYPE_STRING);
                summary = rows.getCell(7).getStringCellValue();
            }
            if (rows.getCell(8) != null){
                rows.getCell(8).setCellType(Cell.CELL_TYPE_STRING);
                ttLCMark = rows.getCell(8).getStringCellValue();
            }
            if (rows.getCell(9) != null){
                rows.getCell(9).setCellType(Cell.CELL_TYPE_STRING);
                customerCode = rows.getCell(9).getStringCellValue();
            }
            if (rows.getCell(10) != null){
                rows.getCell(10).setCellType(Cell.CELL_TYPE_STRING);
                customerName = rows.getCell(10).getStringCellValue();
            }
            if (rows.getCell(11) != null){
                rows.getCell(11).setCellType(Cell.CELL_TYPE_STRING);
                recognizedAmount = rows.getCell(11).getStringCellValue();
            }
            if (rows.getCell(12) != null){
                rows.getCell(12).setCellType(Cell.CELL_TYPE_STRING);
                productCode = rows.getCell(12).getStringCellValue();
            }
            if (rows.getCell(13) != null){
                rows.getCell(13).setCellType(Cell.CELL_TYPE_STRING);
                pi = rows.getCell(13).getStringCellValue();
            }
            if (rows.getCell(14) != null){
                rows.getCell(14).setCellType(Cell.CELL_TYPE_STRING);
                prepayment = rows.getCell(14).getStringCellValue();
            }
            if (rows.getCell(15) != null){
                rows.getCell(15).setCellType(Cell.CELL_TYPE_STRING);
                transferTo = rows.getCell(15).getStringCellValue();
            }else {
                transferTo = "";
            }
            if (rows.getCell(16) != null){
                rows.getCell(16).setCellType(Cell.CELL_TYPE_STRING);
                depositPrincipal = rows.getCell(16).getStringCellValue();
            }
            if (rows.getCell(17) != null){
                rows.getCell(17).setCellType(Cell.CELL_TYPE_STRING);
                depositInterest = rows.getCell(17).getStringCellValue();
            }
            if (rows.getCell(18) != null){
                rows.getCell(18).setCellType(Cell.CELL_TYPE_STRING);
                staffName = rows.getCell(18).getStringCellValue();
            }
            if (rows.getCell(18) != null){
                rows.getCell(18).setCellType(Cell.CELL_TYPE_STRING);
                remark = rows.getCell(18).getStringCellValue();
            }
            if (rows.getCell(0) != null){
                rows.getCell(0).setCellType(Cell.CELL_TYPE_STRING);
                sapIncomeNo = rows.getCell(0).getStringCellValue();
            }
            if (rows.getCell(20) != null){
                rows.getCell(20).setCellType(Cell.CELL_TYPE_STRING);
                comment = rows.getCell(20).getStringCellValue();
            }
            if (rows.getCell(21) != null){
                rows.getCell(21).setCellType(Cell.CELL_TYPE_STRING);
                sapNoForOther = rows.getCell(21).getStringCellValue();
            }
            if (rows.getCell(22) != null){
                rows.getCell(22).setCellType(Cell.CELL_TYPE_STRING);
                sapClearingNo = rows.getCell(22).getStringCellValue();
            }
            if (rows.getCell(23) != null){
                rows.getCell(23).setCellType(Cell.CELL_TYPE_STRING);
                EmailAddress = rows.getCell(23).getStringCellValue();
            }
            MiddleEastAll middleEastAll = new MiddleEastAll();
            middleEastAll.setId(id);
            middleEastAll.setBank(bank);
            middleEastAll.setDocumentDate(documentDate);
            middleEastAll.setCurrency(currency);
            middleEastAll.setIncome(income);
            middleEastAll.setCharge(charge);
            middleEastAll.setSummary(summary);
            middleEastAll.setTtLCMark(ttLCMark);
            middleEastAll.setCustomerCode(customerCode);
            middleEastAll.setCustomerName(customerName);
            middleEastAll.setRecognizedAmount(recognizedAmount);
            middleEastAll.setPrepayment(prepayment);
            middleEastAll.setProductCode(productCode);
            middleEastAll.setPi(pi);
            middleEastAll.setTransferTo(transferTo);
            middleEastAll.setDepositPrincipal(depositPrincipal);
            middleEastAll.setDepositInterest(depositInterest);
            middleEastAll.setStaffName(staffName);
            middleEastAll.setSapIncomeNo(sapIncomeNo);
            middleEastAll.setRemark(remark);
            middleEastAll.setComment(comment);
            middleEastAll.setSapNoForOther(sapNoForOther);
            middleEastAll.setSapClearingNo(sapClearingNo);
            middleEastAll.setEmail(EmailAddress);
            if (!"".equals(middleEastAll.getId())){
                MiddleEastList.add(middleEastAll);
            }
        }
        return MiddleEastList;
    }

    /**
     * 数据处理
     * 合并单元格数据处理
     * @param MiddleEastList
     * @return
     */
    public static List HandleData(List<MiddleEastAll> MiddleEastList){
        for (int i = 1 ;i < MiddleEastList.size()-1;i++ ){
            MiddleEastAll middleEastAll = MiddleEastList.get(i);
            MiddleEastAll middleEastAll2 = MiddleEastList.get(i - 1);
            if ("".equals(middleEastAll.getId()) || middleEastAll.getId().equals(middleEastAll2.getId())){
                middleEastAll.setId(middleEastAll2.getId());
                middleEastAll.setBank(middleEastAll2.getBank());
                middleEastAll.setDocumentDate(middleEastAll2.getDocumentDate());
                middleEastAll.setIncome(middleEastAll2.getIncome());
                middleEastAll.setCharge(middleEastAll2.getCharge());
                middleEastAll.setCurrency(middleEastAll2.getCurrency());
                middleEastAll.setSummary(middleEastAll2.getSummary());
                middleEastAll.setTtLCMark(middleEastAll2.getTtLCMark());
                middleEastAll.setCustomerCode(middleEastAll2.getCustomerCode());
                middleEastAll.setCustomerName(middleEastAll2.getCustomerName());
            }
        }
        return MiddleEastList;
    }

    /**
     * 中东区AR核销数据处理生成台账表
     * @param MiddleEastALLList
     */
    public static void excelOutput_Log(List<MiddleEastAll> MiddleEastALLList, String filePath) throws Exception{
        //创建表格
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        //定义第一个sheet页
        XSSFSheet xssfSheet = xssfWorkbook.createSheet("Sheet1");
        //第一个sheet页数据（生成台账表）
        Row row0 = xssfSheet.createRow(0);
        String[] headers = new String[]{"SAP Income No.","ID","Bank", "Document Date", "Income","Charge", "Currency", "Summary","TTLCMark", "Customer Code", "Customer Name", "Recognized Amount","Product Code","PI","Prepayment", "Transfer To", "Deposit-Principal", "Deposit-Interest","Staff Name","Remark","Comment","SAP No. for Other","SAP Clearing No.","EmailAddress"};
        for (int i = 0; i < headers.length; i++) {
            XSSFCell cell = (XSSFCell) row0.createCell(i);
            xssfSheet.setColumnWidth(i + 2, 5000);
            XSSFRichTextString text = new XSSFRichTextString(headers[i]);
            cell.setCellValue(text);
        }
        int rowNum = 1;
        for (MiddleEastAll middleEastAll : MiddleEastALLList){
            XSSFRow row = xssfSheet.createRow(rowNum);
            row.createCell(1).setCellValue(middleEastAll.getId());
            row.createCell(2).setCellValue(middleEastAll.getBank());
            row.createCell(3).setCellValue(middleEastAll.getDocumentDate());
            if (middleEastAll.getIncome().contains(".")){
                row.createCell(4).setCellValue(convertData(middleEastAll.getIncome()));
            }else {
                row.createCell(4).setCellValue(middleEastAll.getIncome());
            }
            if (middleEastAll.getCharge().contains(".")){
                row.createCell(5).setCellValue(convertData(middleEastAll.getCharge()));
            }else {
                row.createCell(5).setCellValue(middleEastAll.getCharge());
            }
            row.createCell(6).setCellValue(middleEastAll.getCurrency());
            row.createCell(7).setCellValue(middleEastAll.getSummary());
            row.createCell(8).setCellValue(middleEastAll.getTtLCMark());
            row.createCell(9).setCellValue(middleEastAll.getCustomerCode());
            row.createCell(10).setCellValue(middleEastAll.getCustomerName());
            if (middleEastAll.getRecognizedAmount() != null){
                if (middleEastAll.getRecognizedAmount().contains(".")){
                    row.createCell(11).setCellValue(convertData(middleEastAll.getRecognizedAmount()));
                }else {
                    row.createCell(11).setCellValue(middleEastAll.getRecognizedAmount());
                }
            }else {
                row.createCell(11).setCellValue(middleEastAll.getRecognizedAmount());
            }
            row.createCell(12).setCellValue(middleEastAll.getProductCode());
            row.createCell(13).setCellValue(middleEastAll.getPi());
            row.createCell(14).setCellValue(middleEastAll.getPrepayment());
            row.createCell(15).setCellValue(middleEastAll.getTransferTo());
            row.createCell(16).setCellValue(middleEastAll.getDepositPrincipal());
            row.createCell(17).setCellValue(middleEastAll.getDepositInterest());
            row.createCell(18).setCellValue(middleEastAll.getStaffName());
            row.createCell(19).setCellValue(middleEastAll.getRemark());
            row.createCell(0).setCellValue(middleEastAll.getSapIncomeNo());
            row.createCell(20).setCellValue(middleEastAll.getComment());
            row.createCell(21).setCellValue(middleEastAll.getSapNoForOther());
            row.createCell(22).setCellValue(middleEastAll.getSapClearingNo());
            row.createCell(23).setCellValue(middleEastAll.getEmail());
            rowNum++;
        }

        //获取业务邮件数据
        List<MiddleEastAll> MiddleEastList = getAddressExcelData(filePath);
        //定义第一个sheet页(发票号任务表)
        XSSFSheet xssfSheet2 = xssfWorkbook.createSheet("Sheet2");
        //第二个sheet页数据
        Row row = xssfSheet2.createRow(0);
        String[] headers1 = new String[]{"Address"};
        for (int i = 0; i < headers1.length; i++) {
            XSSFCell cell = (XSSFCell) row.createCell(i);
            xssfSheet2.setColumnWidth(i, 5000);
            XSSFRichTextString text = new XSSFRichTextString(headers1[i]);
            cell.setCellValue(text);
        }
        int rowNum1 = 1;
        for (MiddleEastAll middleEastAll : MiddleEastList){
            XSSFRow rows = xssfSheet2.createRow(rowNum1);
            xssfSheet2.setColumnWidth(rowNum1, 5000);
            rows.createCell(0).setCellValue(middleEastAll.getRemark());;
            rowNum1++;
        }

        FileOutputStream fileOutputStream = new FileOutputStream(filePath);
        xssfWorkbook.write(fileOutputStream);
    }

    /**
     * 获取业务邮件的数据
     * @param filePath
     * @return
     */
    public static List getAddressExcelData(String filePath){
        List<MiddleEastAll> MiddleEastList = new ArrayList<>();
        try {
            String address = "";
            File excelFile = new File(filePath);
            FileInputStream EIL_file_IO = new FileInputStream(excelFile);
            Workbook wb = WorkbookFactory.create(EIL_file_IO);
            Sheet sheet = wb.getSheetAt(1);
            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row rows = sheet.getRow(r);//获取第r行
                if (rows.getCell(0) != null){
                    rows.getCell(0).setCellType(Cell.CELL_TYPE_STRING);
                    address = rows.getCell(0).getStringCellValue();
                }else {
                    address = "";
                }
                MiddleEastAll middleEastAll = new MiddleEastAll();
                middleEastAll.setRemark(address);
                if (!"".equals(address)){
                    MiddleEastList.add(middleEastAll);
                }
            }
        }catch (Exception e){
            MiddleEastAll middleEastAll = new MiddleEastAll();
            MiddleEastList.add(middleEastAll);
        }
        return MiddleEastList;
    }

}
