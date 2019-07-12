package HttpUtil;

import DataClean.CustomerCode;
import DataClean.Ledger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 俄罗斯AR核销数据处理
 */
public class Russia_AR {
    public static void main(String[] args) {
        String filePath = args[0];//原始数据表
        String filePath_cus = args[1];//对照表
        String filePath_ven = args[2];//对照表
        String fileName = args[3];//生成台账表格路径
//        String filePath = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\俄罗斯AR2\\test\\0416-101.xlsx";//原始数据表
//        String filePath_cus = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\俄罗斯AR2\\test\\62F0CustomerList.xlsx";//对照表
//        String filePath_ven = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\俄罗斯AR\\62F0vendorlist.xlsx";//对照表
//        String fileName = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\俄罗斯AR2\\test\\101_basicinfo.xlsx";//生成台账表格路径

        File excelFile = new File(filePath);
        String code = filePath.substring(filePath.lastIndexOf("-")+1,filePath.lastIndexOf("."));
        System.out.println(code);

        File excelFile_cus = new File(filePath_cus);
        File excelFile_ven = new File(filePath_ven);
        try {
            getLedgerExcel(excelFile, excelFile_cus, excelFile_ven, fileName,code);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 数据处理
     * @param excelFile
     * @param excelFile_cus
     * @param excelFile_ven
     * @param fileName
     * @throws Exception
     */
    public static void getLedgerExcel( File excelFile, File excelFile_cus, File excelFile_ven, String fileName,String typeCode) throws Exception{
        //获取原始数据
        FileInputStream EIL_file_IO = new FileInputStream(excelFile);
        Workbook wb =  WorkbookFactory.create(EIL_file_IO);
        Sheet sheet = wb.getSheetAt(0);
        //获取生成台账数据
        List<Ledger> LedList =  getLedgerLise(sheet);
        System.out.println("输出台账条数：" + LedList.size());
        //根据code写GL
        for (Ledger ledger:LedList){
           switch (typeCode){
               case "101":
                   ledger.setGL("1002030710");
                   break;
               case "102":
                   ledger.setGL("1002030270");
                   break;
               case "103":
                   ledger.setGL("1002030280");
                   break;
               case "104":
                   ledger.setGL("1002030360");
                   break;
               case "106":
                   ledger.setGL("1002030720");
                   break;
               default:
                   ledger.setGL("");
                   break;
           }
        }
        //获取Customer对照表数据
        FileInputStream EIL_file_IO1 = new FileInputStream(excelFile_cus);
        Workbook wb_cus =  WorkbookFactory.create(EIL_file_IO1);
        Sheet sheet_cus = wb_cus.getSheetAt(0);
        List<CustomerCode> CustomerList = getCustomerCode(sheet_cus);//调用方法获取所需数据

        //获取vendor对照表数据
        FileInputStream EIL_file_IO_ven = new FileInputStream(excelFile_ven);
        Workbook wb_ven =  WorkbookFactory.create(EIL_file_IO_ven);
        Sheet sheet_ven = wb_ven.getSheetAt(0);
        List<CustomerCode> Customer_venList = getCustomer_venCode(sheet_ven);//调用方法获取所需数据

        //补存缺失字段.以及提取有效的单号
        Map<String ,Object> map = handleDate(LedList);
        LedList = (List<Ledger>) map.get("Ledger");//补存字段之后返回数据
        //通过对照表添加CustomerCode
        //首先判断台账数据的ReasonCode是否为184，若是，对照vendor表数据查询CustomerCode
        for (Ledger ledger : LedList)
            if ("184".equals(ledger.getReasonCode())) {
                for (CustomerCode code : Customer_venList) {
                    if (code.getTaxNum().equals(ledger.getTaxCode())) {
                        ledger.setCustomerCode(code.getCustomerCode());
                    }
                }
            } else {//若ReasonCode不是184，则对照customer表数据查询CustomerCode
                //用来存储税号（taxNum）重复的数据
                List<CustomerCode> cusCode = new ArrayList<>();
                //查看customer数据表，将表中taxNum值与台账表TaxCode值相等的数据都取出来,（taxNum可能会有重复值）
                for (CustomerCode codes : CustomerList) {
                    if (codes.getTaxNum().equals(ledger.getTaxCode())) {
                        cusCode.add(codes);
                    }
                }
                //判断是否有税号重复数据
                if (cusCode.size() > 1) {
                    //用来存储在税号相同的情况下，name值也相同
                    List<CustomerCode> cusCode1 = new ArrayList<>();
                    for (CustomerCode code1 : cusCode) {
                        //若有多条数据满足条件，则在判断Customer
//                        System.out.println("ledger.getCustomer()："+ledger.getCustomer());
//                        System.out.println("code1:"+code1);
//                        System.out.println("code1.getName()："+code1.getName());
                        if (code1.getName()!=null){
                            if (ledger.getCustomer().contains(code1.getName())) {
                                cusCode1.add(code1);
                            }
                        }
                    }
                    //判断name是否唯一
                    if (cusCode1.size() == 1){
                        for (CustomerCode code2 : cusCode1) {
                            ledger.setCustomerCode(code2.getCustomerCode());
                            //通过判断对照表中国家标识，填充台账中国家字段
                            if (code2.getCountry() != null && "RU".equals(code2.getCountry())) {
                                ledger.setCountry("XC");
                            }
                            if (code2.getCountry() != null && !"RU".equals(code2.getCountry())) {
                                ledger.setCountry("AA");
                            }
                        }
                    }
                    continue;
                }
                //若customer表数据taxNum值，与台账表TaxCode值唯一相等，则直接提取有效数据
                if (cusCode.size() == 1) {
                    for (CustomerCode code : cusCode) {
                        ledger.setCustomerCode(code.getCustomerCode());
                        if (code.getCountry() != null && "RU".equals(code.getCountry())) {
                            ledger.setCountry("XC");
                        }
                        if (code.getCountry() != null && !"RU".equals(code.getCountry())) {
                            ledger.setCountry("AA");
                        }
                    }
                    continue;
                }
                //最后判断CustomerCode是否为null,是则使用name与Customer查询满足条件数据
                if (ledger.getCustomerCode() == null){
                    List<CustomerCode> cusCode2 = new ArrayList<>();
                    for (CustomerCode code : CustomerList) {
                        if (code.getName() != null && ledger.getCustomer().contains(code.getName())) {
                            cusCode2.add(code);
                        }
                    }
                    if (cusCode2.size() == 1){
                        for (CustomerCode code : cusCode2) {
                            ledger.setCustomerCode(code.getCustomerCode());
                            if (code.getCountry() != null && "RU".equals(code.getCountry())) {
                                ledger.setCountry("XC");
                            }
                            if (code.getCountry() != null && !"RU".equals(code.getCountry())) {
                                ledger.setCountry("AA");
                            }
                        }
                    }
                }
            }
        //获取发票任务表数据
        List<CustomerCode> codeList = (List<CustomerCode>) map.get("CustomerCode");
        List<CustomerCode> codeList1 = new ArrayList<>();
        if (codeList.size() > 0 && LedList.size() > 0){
            for (CustomerCode code : codeList){
                List<String> invoiceNoList = code.getInvoiceNoList();//获取每条数据对应的单号的集合
                for (String invoice : invoiceNoList){
                    CustomerCode code1 = new CustomerCode();
                    code1.setIndex(code.getIndex());
                    code1.setCustomerCode(code.getCustomerCode());
                    code1.setInvoiceNo(invoice);
                    code1.settCode(code.gettCode());
                    code1.setText(code.getText());
                    code1.setAmount(code.getAmount());
                    code1.setDate(code.getDate());
                    codeList1.add(code1);
                }
            }
            //发票任务表数据补存CustomerCode
            for (int i = 0; i < LedList.size() ;i ++){
                for (CustomerCode code :codeList1 ){
                    if (String.valueOf( i + 1).equals(code.getIndex())){
                        code.setCustomerCode(LedList.get(i).getCustomerCode());
                    }
                }
            }
        }

//        customerCode 为 6000240489的数sapClear为No Need
            for (Ledger ledger:LedList){
                if (ledger.getCustomerCode()!=null){
                    if (ledger.getCustomerCode().equals("6000240489")){
                        ledger.setSapClearNo("NO NEED");
                    }
                }
            }
            //生成台账表以及发票任务表
        excelOutput_Excel(LedList,fileName,codeList1);
        System.out.println("运行结束！" );
    }

    /**
     * 提取有效数字
     * @param data
     * @return
     */
    public static String clearData(String data){
        Pattern p = Pattern.compile("[^.0-9]");//提取有效数字
        data = p.matcher(data).replaceAll("").trim();
        return data;
    }
    /**
     * 获取原始台账数据
     * @param sheet
     * @return
     */
    public static List getLedgerLise(Sheet sheet){
        String date = "";
        String customer = "";
        String title = "";
        String taxCode = "";
        String amount = "";
        DateFormat dFormat = new SimpleDateFormat("MM/dd/yyyy");
        //遍历获取到的表格数据，提取需要的行、列
        List<Ledger> LedList = new ArrayList<>();
        for (int r = 10; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            Ledger ledger = new Ledger();
            if (row.getCell(12) != null){
                row.getCell(12).setCellType(Cell.CELL_TYPE_STRING);
                String am = clearData(row.getCell(12).getStringCellValue());
                BigDecimal bd = new BigDecimal(am);
                amount = bd.setScale(2, BigDecimal.ROUND_HALF_UP).toPlainString();
                ledger.setAmount(amount);
            }else {
                amount = "";
            }
            if (!"".equals(amount) && Double.parseDouble(amount) != 0){//判断amount是否为0
                if (row.getCell(0) != null){
                    row.getCell(0).setCellType(Cell.CELL_TYPE_STRING);
                    date = clearData(row.getCell(0).getStringCellValue());
                    date = date.replace(".","/");
                    date = date.substring(3,6) + date.substring(0,3) + date.substring(6);
                    ledger.setDate(date);
                }
                if (row.getCell(2) != null){
                    row.getCell(2).setCellType(Cell.CELL_TYPE_STRING);
                    title = row.getCell(2).getStringCellValue();
                    ledger.setTitle(title);
                }
                if (row.getCell(5) != null){
                    row.getCell(5).setCellType(Cell.CELL_TYPE_STRING);
                    customer = row.getCell(5).getStringCellValue();
                        ledger.setCustomer(customer);
                }
                if (row.getCell(8) != null){
                    row.getCell(8).setCellType(Cell.CELL_TYPE_STRING);
                    String original_taxcode = row.getCell(8).getStringCellValue();
                    taxCode = clearData(row.getCell(8).getStringCellValue());
                    if (!taxCode.equals("")){
                        taxCode = new BigDecimal(taxCode).toString();
                    }else {
                        taxCode = "";
                    }
                   // taxCode = bd.setScale(0, BigDecimal.ROUND_HALF_UP).toPlainString();
                    ledger.setTaxCode(taxCode);
                }
                if (!customer.contains("ООО \"Компания РБТ\"")){
                    LedList.add(ledger);
                }
            }else {
                break;
            }
        }
        return LedList;
    }
    /**
     * 获取对照表格的数据
     * @param sheet
     * @return
     */
    public static List getCustomerCode(Sheet sheet){
        String customerCode = "";
        String country = "";
        String taxNum = "";
        String name = "";
        List<CustomerCode> CustomerCodeList = new ArrayList<>();
        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            CustomerCode customer = new CustomerCode();
            if (row.getCell(0) != null){
                row.getCell(0).setCellType(Cell.CELL_TYPE_STRING);
                customerCode = row.getCell(0).getStringCellValue();
                customer.setCustomerCode(customerCode);
            }
            if (row.getCell(1) != null){
                row.getCell(1).setCellType(Cell.CELL_TYPE_STRING);
                name = row.getCell(1).getStringCellValue();
                customer.setName(name);
            }
            if (row.getCell(11) != null){
                row.getCell(11).setCellType(Cell.CELL_TYPE_STRING);
                country = row.getCell(11).getStringCellValue();
                customer.setCountry(country);
            }
            if (row.getCell(18) != null) {
                row.getCell(18).setCellType(Cell.CELL_TYPE_STRING);
                taxNum = clearData(row.getCell(18).getStringCellValue());
                customer.setTaxNum(taxNum);
            }
            CustomerCodeList.add(customer);
        }
        return CustomerCodeList;
    }

    /**
     * 获取vendor对照表格的数据
     * @param sheet
     * @return
     */
    public static List getCustomer_venCode(Sheet sheet){
        String customerCode = "";
        String taxNum = "";
        List<CustomerCode> CustomerCodeList = new ArrayList<>();
        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            CustomerCode customer = new CustomerCode();
            if (row.getCell(0) != null){
                row.getCell(0).setCellType(Cell.CELL_TYPE_STRING);
                customerCode = row.getCell(0).getStringCellValue();
                customer.setCustomerCode(customerCode);
            }
            if (row.getCell(37) != null) {
                row.getCell(37).setCellType(Cell.CELL_TYPE_STRING);
                taxNum = clearData(row.getCell(37).getStringCellValue());
                customer.setTaxNum(taxNum);
            }
            CustomerCodeList.add(customer);
        }
        return CustomerCodeList;
    }

    /**
     * 补存缺失数据（填写ReasonCode、SapClearNo）
     * 并返回最后结果数据，以及发票号任务表数据
     * 该方法通过三个等级的判断，完成数据补存
     * 并获取每天数据的所有invoiceNo
     * @param list
     * @return
     */
    public static Map handleDate(List<Ledger> list){
        Map<String,Object> map = new HashMap<>();
        List<CustomerCode> codeList = new ArrayList<>();
        //遍历集合
        for (int i = 0 ; i < list.size(); i++){
            CustomerCode customerCode = new CustomerCode();
            Ledger ledger = list.get(i);
            if (ledger.getCustomer().contains("ООО \"ИМПЕРИАЛ\"") || ledger.getCustomer().contains("ООО Курьер-Регион Столица")){
                ledger.setSapClearNo("NO NEED");
            }
            if (ledger.getCustomer().contains("ООО \"ЕВРАЗИЯ ЛОДЖИСТИК\"")){
                ledger.setCustomerCode("6000240806");
                ledger.setReasonCode("102");
                ledger.setSapClearNo("NO NEED");
                ledger.setText("Перечисление денежных средств");

            }
            String title = ledger.getTitle();
            //判断title中是否包含指定内容--并设定ReasonCode
            if (title.contains("Возврат депозита по заявке")){
                ledger.setReasonCode("100");
                ledger.setCustomerCode("1002030740");
                ledger.setText("Возврат депозита по заявке \"ХАР\" ООО");
                ledger.setSapClearNo("本金");
                continue;
            }
            if (title.contains("Выплата процентов по депозиту")){
                ledger.setReasonCode("332");
                ledger.setCustomerCode("5503010000");
                ledger.setText(title.substring(1,title.indexOf("\"ХАР\"")));
                ledger.setSapClearNo("利息");
                continue;
            }
            if (title.contains("Возврат излишне перечисленных денеж-ых средств") || title.contains("Оплата по претензиям")){
                ledger.setReasonCode("184");
                ledger.setSapClearNo("供应商返还款");
                continue;
            }
            String reasonCode = ledger.getReasonCode();
            String customer = ledger.getCustomer();
            //若ReasonCode为null，继续操作
            if (reasonCode == null){
                if (title.contains("Пер-е")){
                    ledger.setReasonCode("102");
                    String str = title.substring(1,title.indexOf("("));
                    ledger.setText(str);
                    ledger.setSapClearNo("NO NEED");
                    continue;
                }
                if (customer.contains("ПАО РОСБАНК")){
                    ledger.setReasonCode("102");
                    String str = title.substring(1,title.indexOf(","));
                    ledger.setText(str);
                    ledger.setSapClearNo("NO NEED");
                    continue;
                }
                if (customer.contains("ООО Курьер-Регион Столица")){
                    ledger.setReasonCode("102");
                    ledger.setSapClearNo("NO NEED");
                    continue;
                }
            }
            String reasonCode1 = ledger.getReasonCode();
            //若ReasonCode为null，继续操作
            if (reasonCode1 == null){
                System.out.println("title:"+title);

                String[] vl = title.split("( |ф)");
                System.out.println(Arrays.toString(vl));
                List<String> list1 = Arrays.asList(vl);
                List<String> list2 = new ArrayList<>();
                for (String s : list1){
                    Pattern pattern = Pattern.compile(".*\\d+.*");
                    Matcher m = pattern.matcher(s);
                    if (m.matches()) {//判断是否有数字
//                        System.out.println("shuchu :" + s);
                        if (s.length() >= 7){
                            Pattern p = Pattern.compile("[^-,.;/0-9]");//提取数字字符，以及不可去除的符号
                            s = p.matcher(s).replaceAll("").trim();
                            if (String.valueOf(s.charAt(0)).equals(".") || String.valueOf(s.charAt(0)).equals(",")||String.valueOf(s.charAt(0)).equals("-")||String.valueOf(s.charAt(0)).equals(";")){
                                s = s.substring(1);
                            }
                            list2.add(s);
                        }
                    }
                }

                //遍历集合
                if (list2.size() > 0){
                    List<String> list3 = new ArrayList<>();//创建集合用来存储提取出来的单号
                    for (String s : list2){

//                        System.out.println("刚开始："+s);

                        //System.out.println("输出每一条数据：" + s);
                        //Title中发票号90开头8位数，ReasonCode用112，
//                        if (s.contains("90")){
                        if (s.indexOf("90",0) == 0){
                            Pattern p = Pattern.compile("[^,/0-9]");//提取有效数字
                            s = p.matcher(s).replaceAll("").trim();
//                            System.out.println(s);

                                if (!s.contains(",") && s.length() == 8){
                                    list3.add(s);
                                    ledger.setReasonCode("112");
                                    customerCode.settCode("FBL5N-9");
                                }else {
                                    String[] vls = s.split(",");
                                    List<String> lists = Arrays.asList(vls);
                                    for (String ss : lists){
                                        System.out.println(ss);
                                        if (s.indexOf("90",0) == 0 && ss.length() == 8){
                                            list3.add(ss);
                                            ledger.setReasonCode("112");
                                            customerCode.settCode("FBL5N-9");
                                        }
                                    }
                                }
                            }
//                        }

                        //Title中发票号0开头8位数，ReasonCode用107，
                        if (s.contains("/")){

                            Pattern p = Pattern.compile("[^-,;/0-9]");//提取有效数字
                            s = p.matcher(s).replaceAll("").trim();
                            if (!s.contains(";") && !s.contains(",") && s.indexOf("0",0) == 0  && s.indexOf("/",0) == 6){
                                s = s.substring(0,8);
                                list3.add(s);
                                ledger.setReasonCode("107");
                                customerCode.settCode("FBL5N-0");
                            }else {
                                if (s.contains(";")){
                                    String[] vls = s.split(";");
                                    List<String> lists = Arrays.asList(vls);
                                    for (String s1 : lists){
                                        if (s1.indexOf("/",0) == 6 && s1.indexOf("0",0) == 0){
                                            s1 = s1.substring(0,8);
                                            list3.add(s1);
                                            ledger.setReasonCode("107");
                                            customerCode.settCode("FBL5N-0");
                                        }
                                    }
                                }else if (s.contains(",")) {
                                    String[] vls = s.split(",");

                                    List<String> lists = Arrays.asList(vls);
                                    for (String s1 : lists){
                                        if (s1.indexOf("/",0) == 6 && s1.indexOf("0",0) == 0){
//                                            System.out.println(s1);
//                                            如果不够8位，则在/后补1
                                            if (s1.length()<8){
                                                s1 = s1+"1";
                                            }
                                            s1 = s1.substring(0,8);
                                            list3.add(s1);
                                            ledger.setReasonCode("107");
                                            customerCode.settCode("FBL5N-0");
                                        }
                                    }
                                }
                            }
                        }
//                        System.out.println("中间的:："+s);

                        //Title中订单号3开头7位数 用112
                        if (s.indexOf("3",0) == 0 && s.length() >= 7){
                            //判断字符串中是否有“，”，并且长度为7
                            if (!s.contains(",") && s.length() == 7){
                                //判断字符串是否都为数字
                                Pattern pattern = Pattern.compile("[0-9]{1,}");
                                Matcher matcher = pattern.matcher((CharSequence) s);
                                if (matcher.matches()){
                                    list3.add(s);
                                    ledger.setReasonCode("112");
                                    customerCode.settCode("VA03");
                                }
                            }else {//字符串中有“，”
                                //判断“，”第一次出现的位置和最后一次出现的位置
                                if (s.indexOf(",",0) == s.lastIndexOf(",")){
                                    s = s.replace(",","");
                                    if (s.length() == 7 && s.indexOf("3",0) == 0){
                                        //判断字符串是否都为数字
                                        Pattern pattern = Pattern.compile("[0-9]{1,}");
                                        Matcher matcher = pattern.matcher((CharSequence) s);
                                        if (matcher.matches()){
                                            list3.add(s);
                                            ledger.setReasonCode("112");
                                            customerCode.settCode("VA03");
                                        }
                                    }
                                }else {//若“，”出现多次
                                    String[] vls = s.split(",");
                                    List<String> lists = Arrays.asList(vls);
                                    for (String s1 : lists){
                                        if (s1.indexOf("3",0) == 0 && s1.length() == 7){
                                            //判断字符串是否都为数字
                                            Pattern pattern = Pattern.compile("[0-9]{1,}");
                                            Matcher matcher = pattern.matcher((CharSequence) s1);
                                            if (matcher.matches()){
                                                list3.add(s1);
                                                ledger.setReasonCode("112");
                                                customerCode.settCode("VA03");
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        //Title中订单号3开头7位数 用112
                        if (s.indexOf("8",0) == 0 && s.length() >= 8){
                            if (!s.contains(",")){
                                //判断字符串是否都为数字
                                Pattern pattern = Pattern.compile("[0-9]{1,}");
                                Matcher matcher = pattern.matcher((CharSequence) s);
                                if (matcher.matches()){
                                    list3.add(s);
                                    ledger.setReasonCode("112");
                                    customerCode.settCode("VL02N");
                                }
                            }else {
                                if (s.indexOf(",",0) == s.lastIndexOf(",")){
                                    s = s.replace(",","");
                                    if (s.length() == 8 && s.indexOf("8",0) == 0){
                                        //判断字符串是否都为数字
                                        Pattern pattern = Pattern.compile("[0-9]{1,}");
                                        Matcher matcher = pattern.matcher((CharSequence) s);
                                        if (matcher.matches()){
                                            list3.add(s);
                                            ledger.setReasonCode("112");
                                            customerCode.settCode("VL02N");
                                        }
                                    }
                                }else {
                                    String[] vls = s.split(",");
                                    List<String> lists = Arrays.asList(vls);
                                    for (String s1 : lists){
                                        if (s1.indexOf("8",0) == 0 && s1.length() == 8){
                                            //判断字符串是否都为数字
                                            Pattern pattern = Pattern.compile("[0-9]{1,}");
                                            Matcher matcher = pattern.matcher((CharSequence) s1);
                                            if (matcher.matches()){
                                                list3.add(s1);
                                                list3.add(s);
                                                ledger.setReasonCode("112");
                                                customerCode.settCode("VL02N");
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    String str = "";
                    //若集合中有数据，处理集合数据，
                    System.out.println(list3.size());
                    if (list3.size() > 0){
                        customerCode.setIndex(String.valueOf(i + 1));
                        customerCode.setInvoiceNoList(list3);
                        customerCode.setDate(ledger.getDate());
                        customerCode.setAmount(ledger.getAmount());
                        String text = "";//将发票号处理并拼接
                        for (int j = 0; j <list3.size(); j++){
                            text = list3.get(j);
                            if (text.contains("90") && text.indexOf("90",0) == 0){
                                if (j > 0){
                                    text = text.substring(3);
                                }
                            }
                            if (text.indexOf("0",0) == 0 && text.indexOf("/",0) == 6){
                                if (j > 0 ){
                                    text = text.substring(1,6);
                                }
                            }
                            str = str + text + "#";
                            ledger.setText(str.substring(0,str.length() - 1));
                        }
                        customerCode.setText(ledger.getText());
                        codeList.add(customerCode);
                    }
                }
            }
//            System.out.println("##############################");
            String reasonCode2 = ledger.getReasonCode();
            if (reasonCode2 == null){
                if (title.contains("деталь") || title.contains("детали") || title.contains("ЗИП") || title.contains("зап.части") || title.contains("запасные части") || title.contains("ЗАПАСНЫЕ ЧАСТИ")
                    || title.contains("ДЕТАЛЬ" )|| title.contains("ДЕТАЛИ") || title.contains("ЗАП.ЧАСТИ") || title.contains("За ТМЦ") || title.contains("за ТМЦ") || title.contains("ЗА ТМЦ")){
                    ledger.setReasonCode("107");
                    continue;
                }
                if (title.contains("оборудование") || title.contains("ОБОРУДОВАНИЕ") || title.contains("ЗА БЫТОВУЮ ТЕХНИКУ") || title.contains("за бытовую технику")){
                    ledger.setReasonCode("112");
                    continue;
                }
            }
        }
        map.put("Ledger",list);
        map.put("CustomerCode",codeList);
        return map;
    }


    /**
     * 生成新的台账表格
     * @param ledgerList
     * @param fileName
     */
    public static void excelOutput_Excel(List<Ledger> ledgerList, String fileName, List<CustomerCode> codeList) throws Exception{
        //创建表格
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        //定义第一个sheet页
        XSSFSheet xssfSheet = xssfWorkbook.createSheet("Sheet1");
        //第一个sheet页数据（生成台账表）
        Row row0 = xssfSheet.createRow(0);
        String[] headers = new String[]{"No","Date","Customer","Title", "Tax Code", "Customer Code","Country", "Amount","Reason Code","Text", "Invoice", "SAPIncomeNo.", "SAPClearNo.", "Status", "ReceivedInfor.Date", "KeyinDate(ClearDate)","Comment","GL","#"};
        for (int i = 0; i < headers.length; i++) {
            xssfSheet.setColumnWidth(i + 1, 5000);
            XSSFCell cell = (XSSFCell) row0.createCell(i);
            XSSFRichTextString text = new XSSFRichTextString(headers[i]);
            cell.setCellValue(text);
        }
        int rowNum = 1;
        for (Ledger ledger : ledgerList){
            XSSFRow row = xssfSheet.createRow(rowNum);
            xssfSheet.setColumnWidth(rowNum, 5000);
            row.createCell(0).setCellValue(rowNum);
            row.createCell(1).setCellValue(ledger.getDate());
            row.createCell(2).setCellValue(ledger.getCustomer());
            row.createCell(3).setCellValue(ledger.getTitle());
            row.createCell(4).setCellValue(ledger.getTaxCode());
            row.createCell(5).setCellValue(ledger.getCustomerCode());
            row.createCell(6).setCellValue(ledger.getCountry());
            row.createCell(7).setCellValue(ledger.getAmount());
            row.createCell(8).setCellValue(ledger.getReasonCode());
            row.createCell(9).setCellValue(ledger.getText());
            row.createCell(10).setCellValue(ledger.getInvoice());
            row.createCell(11).setCellValue(ledger.getSapIncomeNo());
            row.createCell(12).setCellValue(ledger.getSapClearNo());
            row.createCell(13).setCellValue(ledger.getStatus());
            row.createCell(14).setCellValue(ledger.getReceivedInforDate());
            row.createCell(15).setCellValue(ledger.getKeyinDate());
            row.createCell(16).setCellValue(ledger.getComment());
            row.createCell(17).setCellValue(ledger.getGL());
            rowNum++;
        }

        //定义第二个sheet页(发票号任务表)
        XSSFSheet xssfSheet2 = xssfWorkbook.createSheet("Sheet2");
        //第二个sheet页数据
        Row row1 = xssfSheet2.createRow(0);
        String[] headers1 = new String[]{"No","Customer Code","Invoice No.","T-Code","Date","Amount","Text"};
        for (int i = 0; i < headers1.length; i++) {
            xssfSheet2.setColumnWidth(i + 1, 5000);
            XSSFCell cell = (XSSFCell) row1.createCell(i);
            XSSFRichTextString text = new XSSFRichTextString(headers1[i]);
            cell.setCellValue(text);
        }
        int rowNum1 = 1;
        for (CustomerCode customerCode : codeList){
            XSSFRow row = xssfSheet2.createRow(rowNum1);
            xssfSheet2.setColumnWidth(rowNum1, 5000);
            row.createCell(0).setCellValue(customerCode.getIndex());
            row.createCell(1).setCellValue(customerCode.getCustomerCode());
            row.createCell(2).setCellValue(customerCode.getInvoiceNo());
            row.createCell(3).setCellValue(customerCode.gettCode());
            row.createCell(4).setCellValue(customerCode.getDate());
            row.createCell(5).setCellValue(customerCode.getAmount());
            row.createCell(6).setCellValue(customerCode.getText());
            rowNum1++;
        }
        FileOutputStream fileOutputStream = new FileOutputStream(fileName);
        xssfWorkbook.write(fileOutputStream);
    }
}

