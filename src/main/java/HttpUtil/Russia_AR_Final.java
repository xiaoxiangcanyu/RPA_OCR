package HttpUtil;

import DataClean.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static HttpUtil.Russia_AR.excelOutput_Excel;

public class Russia_AR_Final extends BaseUtil {
    public static void main(String[] args) {
        String filePath = args[0];//水单路径
//        101收款对照表
        String file_Income_Path = args[1];
        String file_Income_Other_Path = args[2];
//        101付款对照表
        String file_Charge_Cus_Path = args[3];
        String file_Charge_Ven_Path = args[4];

        String filePath2 =args[5];//102对照表
        String filePath3 = args[6];//103对照表
//
//        String filePath = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\俄罗斯AR\\代办\\1128-101.xlsx";

//        101收款对照表
//        String file_Income_Path = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\俄罗斯AR2\\101匹配表.xlsx";
//        String file_Income_Other_Path = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\俄罗斯AR\\代办\\vendorlist_101.xlsx";
//        101付款对照表
//        String file_Charge_Cus_Path = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\俄罗斯AR\\62F0customerlist.xlsx";//对照表
//        String file_Charge_Ven_Path = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\俄罗斯AR\\62F0vendorlist.xlsx";//对照表
//        102
//        String filePath2 = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\俄罗斯AR2\\102扣款vendor_list匹配表.xlsx";
//        103
//        String filePath3 = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\俄罗斯AR2\\103账户客户编码匹配.xlsx";

        String code = filePath.substring(filePath.lastIndexOf("-") + 1, filePath.lastIndexOf("."));
        String filePath_SAP = filePath.substring(0, filePath.lastIndexOf("\\") + 1) + "FB01.xls";
        String file_Path_SAP_101_income = filePath.substring(0, filePath.lastIndexOf("\\") + 1) + "income_FB01.xls";
        String file_Path_SAP_101_charge = filePath.substring(0, filePath.lastIndexOf("\\") + 1) + "charge_FB01.xls";
        String ledger_Path = filePath.substring(0, filePath.lastIndexOf("\\") + 1) + "basicinfo_pay.xlsx";
        try {
            List<RussiaAR> RussiaARList = null;
            List<Russia_AR> RussiaARListPeriod = null;
            switch (code) {
                case "101":
                    // 101 收款
                   getLedgerExcel(new File(filePath), new File(file_Charge_Cus_Path), new File(file_Charge_Ven_Path), file_Path_SAP_101_income,code);
                    // 101 付款
                    List<RussiaLedger> LedList = getLedgerFunction(filePath);
                    if (LedList.size() > 0) {
                        RussiaARList = HandleData(LedList, ledger_Path, file_Income_Path,file_Income_Other_Path);
                    }
                    if (RussiaARList.size() > 0) {
                        excelOutput_FB01(RussiaARList, file_Path_SAP_101_charge);
                    }
                    break;
                case "102":
                    // 102 收款
                    List<RussiaLedger> LedList2 = getLedgerFunction2(filePath);
                    String ledger_Path1 = filePath.substring(0, filePath.lastIndexOf("\\") + 1) + "basicinfo_receive.xlsx";
                    if (LedList2.size() > 0) {
                        RussiaARList = HandleData2(LedList2, ledger_Path1);
                        for (RussiaAR russiaAR : RussiaARList) {
                        }
                    }else {
                        RussiaARList = new ArrayList();
                    }
                    // 102 付款
                    List<RussiaLedger> LedList02 = getLedgerFunction(filePath);
                    if (LedList02.size() > 0) {
                        List<RussiaAR> RussiaARList2 = HandleData02(LedList02, ledger_Path, filePath2);
                        for (RussiaAR russiaAR : RussiaARList2) {

                        }
                        if (RussiaARList2.size() > 0) {
                            RussiaARList.addAll(RussiaARList2);
                        }
                    }
                    break;
                case "103":
                    // 103 收款
                    List<RussiaLedger> LedList3 = getLedgerFunction2(filePath);
                    String ledger_Path3 = filePath.substring(0, filePath.lastIndexOf("\\") + 1) + "basicinfo_receive.xlsx";
                    if (LedList3.size() > 0) {
                        RussiaARList = HandleData3(LedList3, ledger_Path3, filePath3);
                    }
                    break;
                case "106":
                    // 106 付款
                    List<RussiaLedger> LedList6 = getLedgerFunction(filePath);
                    if (LedList6.size() > 0) {
                        RussiaARList = HandleData6(LedList6, ledger_Path);
                    }
                    break;
            }
            //生成FB01表
//            for (RussiaAR russiaAR : RussiaARList) {
//                System.out.println("生成FB01：" + russiaAR.getNo());
//            }
            if (RussiaARList.size() > 0) {
                if (!code.equals("101")){
                    excelOutput_FB01(RussiaARList, filePath_SAP);
                }
            }
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
//        //生成台账表以及发票任务表
        excelOutput_Excel(LedList,fileName,codeList1);
        System.out.println("101收款运行结束！" );
//        return LedList;
    }
    /**
     * 101编号 付款 数据处理
     */
    public static List HandleData(List<RussiaLedger> LedList, String ledger_Path, String banklist,String otherBankList) throws Exception {
        System.out.println("101付款数据处理中！");
        //FB01 数据集合
        List<RussiaAR> RussiaList = new ArrayList<>();
        List<RussiaLedger> LedgerList = new ArrayList<>();
        Set set = new HashSet();
        for (RussiaLedger ledger : LedList) {
            ledger.setGl("1002030710");
            ledger.setSapClearNo("无需清账");
            RussiaAR russiaAR1 = new RussiaAR();
            RussiaAR russiaAR2 = new RussiaAR();
            String title = ledger.getTitle().replace(" ", "");
//            模板-------------》
            if (title.indexOf("Комиссия за валютный контроль") == 0) {
                String str = title.replace("Комиссия за валютный контроль", "").trim();
                String amount = "";
                RussiaLedger ledgers = null;
                for (RussiaLedger ledger1 : LedList) {
                    String title1 = ledger1.getTitle().trim();
                    if (title1.contains(str) && title1.contains("НДС на комиссию за валютный контроль")) {
                        ledgers = ledger1;
                        double sum = Double.parseDouble(clearData(ledger.getAmount())) + Double.parseDouble(clearData(ledger1.getAmount()));
                        BigDecimal bd = new BigDecimal(sum);
                        amount = bd.setScale(2, BigDecimal.ROUND_HALF_UP).toPlainString();
                    }
                }
                if (ledgers != null) {
                    //判断集合中是否已经存在该条数据，不存在则直接添加，存在则删除掉原有数据重新添加
                    if (set.add(ledger.getNo())) {
                        ledger.setReasonCode("362");
                        ledger.setVendorCode("V999995830");
                        LedgerList.add(ledger);
                    } else {
                        int f = LedgerList.indexOf(ledger);
                        LedgerList.remove(f);
                        ledger.setReasonCode("362");
                        ledger.setVendorCode("V999995830");
                        LedgerList.add(ledger);
                    }
                    if (set.add(ledgers.getNo())) {
                        ledgers.setNo(ledger.getNo());
                        ledgers.setReasonCode("362");
                        ledgers.setVendorCode("V999995830");
                        LedgerList.add(ledgers);
                    } else {
                        int f = LedgerList.indexOf(ledgers);
                        LedgerList.remove(f);
                        ledgers.setNo(ledger.getNo());
                        ledgers.setReasonCode("362");
                        ledgers.setVendorCode("V999995830");
                        LedgerList.add(ledgers);
                    }
                    russiaAR1.setNo(ledger.getNo());
                    russiaAR1.setDocumentDate(ledger.getDate());
                    russiaAR1.setType("SA");
                    russiaAR1.setCompanyCode("62F0");
                    russiaAR1.setPostingDate(ledger.getDate());
                    int i = ledger.getDate().indexOf("/", 0);
                    russiaAR1.setPeriod(ledger.getDate().substring(0, i));
                    russiaAR1.setCurrencyRate("RUB");
                    russiaAR1.setPostkey("50");
                    russiaAR1.setGl("1002030710");
                    russiaAR1.setAmount(amount);
                    russiaAR1.setValueDate(ledger.getDate());
                    if (ledger.getTitle().length() > 50) {
                        russiaAR1.setText(ledger.getTitle().substring(0, 50));
                    } else {
                        russiaAR1.setText(ledger.getTitle());
                    }
                    russiaAR1.setReasonCode("362");
                    RussiaList.add(russiaAR1);

                    russiaAR2.setNo(ledger.getNo());
                    russiaAR2.setDocumentDate(ledger.getDate());
                    russiaAR2.setType("SA");
                    russiaAR2.setCompanyCode("62F0");
                    russiaAR2.setPostingDate(ledger.getDate());
                    russiaAR2.setPeriod(ledger.getDate().substring(0, i));
                    russiaAR2.setCurrencyRate("RUB");
                    russiaAR2.setPostkey("25");
                    russiaAR2.setGl("V999995830");
                    russiaAR2.setAmount("*");
                    if (ledgers.getTitle().length() > 50) {
                        russiaAR2.setText(ledgers.getTitle().substring(0, 50));
                    } else {
                        russiaAR2.setText(ledgers.getTitle());
                    }
                    RussiaList.add(russiaAR2);
                }
            }
            if (title.indexOf("Комиссия за выбор банка-посредника согласно тарифам банка") == 0) {
                String str = title.replace("Комиссия за выбор банка-посредника согласно тарифам банка", "").trim();
                String amount = "";
                RussiaLedger ledgers = null;
                for (RussiaLedger ledger1 : LedList) {
                    String title1 = ledger1.getTitle().trim();
                    if (title1.contains(str) && title1.contains("НДС по комиссии за выбор банка-посредника")) {
                        ledgers = ledger1;
                        double sum = Double.parseDouble(clearData(ledger.getAmount())) + Double.parseDouble(clearData(ledger1.getAmount()));
                        BigDecimal bd = new BigDecimal(sum);
                        amount = bd.setScale(2, BigDecimal.ROUND_HALF_UP).toPlainString();
                    }
                }
                if (ledgers != null) {
                    //判断集合中是否已经存在该条数据，不存在则直接添加，存在则删除掉原有数据重新添加
                    if (set.add(ledger.getNo())) {
                        ledger.setReasonCode("362");
                        ledger.setVendorCode("V999995830");
                        LedgerList.add(ledger);
                    } else {
                        int f = LedgerList.indexOf(ledger);
                        LedgerList.remove(f);
                        ledger.setReasonCode("362");
                        ledger.setVendorCode("V999995830");
                        LedgerList.add(ledger);
                    }
                    if (set.add(ledgers.getNo())) {
                        ledgers.setNo(ledger.getNo());
                        ledgers.setReasonCode("362");
                        ledgers.setVendorCode("V999995830");
                        LedgerList.add(ledgers);
                    } else {
                        int f = LedgerList.indexOf(ledgers);
                        LedgerList.remove(f);
                        ledgers.setNo(ledger.getNo());
                        ledgers.setReasonCode("362");
                        ledgers.setVendorCode("V999995830");
                        LedgerList.add(ledgers);
                    }
                    russiaAR1.setNo(ledger.getNo());
                    russiaAR1.setDocumentDate(ledger.getDate());
                    russiaAR1.setType("SA");
                    russiaAR1.setCompanyCode("62F0");
                    russiaAR1.setPostingDate(ledger.getDate());
                    int i = ledger.getDate().indexOf("/", 0);
                    russiaAR1.setPeriod(ledger.getDate().substring(0, i));
                    russiaAR1.setCurrencyRate("RUB");
                    russiaAR1.setPostkey("50");
                    russiaAR1.setGl("1002030710");
                    russiaAR1.setAmount(amount);
                    russiaAR1.setValueDate(ledger.getDate());
                    if (ledger.getTitle().length() > 50) {
                        russiaAR1.setText(ledger.getTitle().substring(0, 50));
                    } else {
                        russiaAR1.setText(ledger.getTitle());
                    }
                    russiaAR1.setReasonCode("362");
                    RussiaList.add(russiaAR1);

                    russiaAR2.setNo(ledger.getNo());
                    russiaAR2.setDocumentDate(ledger.getDate());
                    russiaAR2.setType("SA");
                    russiaAR2.setCompanyCode("62F0");
                    russiaAR2.setPostingDate(ledger.getDate());
                    russiaAR2.setPeriod(ledger.getDate().substring(0, i));
                    russiaAR2.setCurrencyRate("RUB");
                    russiaAR2.setPostkey("25");
                    russiaAR2.setGl("V999995830");
                    russiaAR2.setAmount("*");
                    if (ledgers.getTitle().length() > 50) {
                        russiaAR2.setText(ledgers.getTitle().substring(0, 50));
                    } else {
                        russiaAR2.setText(ledgers.getTitle());
                    }
                    RussiaList.add(russiaAR2);
                }
            }
            if (title.indexOf("КОМИССИЯ") == 0 && !title.contains("валютный контроль")) {
                if (set.add(ledger.getNo())) {
                    ledger.setReasonCode("362");
                    ledger.setVendorCode("5503040000");
                    LedgerList.add(ledger);
                } else {
                    int f = LedgerList.indexOf(ledger);
                    LedgerList.remove(f);
                    ledger.setReasonCode("362");
                    ledger.setVendorCode("5503040000");
                    LedgerList.add(ledger);
                }
                russiaAR1.setNo(ledger.getNo());
                russiaAR1.setDocumentDate(ledger.getDate());
                russiaAR1.setType("SA");
                russiaAR1.setCompanyCode("62F0");
                russiaAR1.setPostingDate(ledger.getDate());
                int i = ledger.getDate().indexOf("/", 0);
                russiaAR1.setPeriod(ledger.getDate().substring(0, i));
                russiaAR1.setCurrencyRate("RUB");
                russiaAR1.setPostkey("50");
                russiaAR1.setGl("1002030710");
                russiaAR1.setAmount(ledger.getAmount());
                russiaAR1.setValueDate(ledger.getDate());
                if (ledger.getTitle().length() > 50) {
                    russiaAR1.setText(ledger.getTitle().substring(0, 50));
                } else {
                    russiaAR1.setText(ledger.getTitle());
                }
                russiaAR1.setReasonCode("362");
                RussiaList.add(russiaAR1);

                russiaAR2.setNo(ledger.getNo());
                russiaAR2.setDocumentDate(ledger.getDate());
                russiaAR2.setType("SA");
                russiaAR2.setCompanyCode("62F0");
                russiaAR2.setPostingDate(ledger.getDate());
                russiaAR2.setPeriod(ledger.getDate().substring(0, i));
                russiaAR2.setCurrencyRate("RUB");
                russiaAR2.setPostkey("40");
                russiaAR2.setGl("5503040000");
                russiaAR2.setAmount("*");
                russiaAR2.setCostCenter("662F120000");
                if (ledger.getTitle().length() > 50) {
                    russiaAR2.setText(ledger.getTitle().substring(0, 50));
                } else {
                    russiaAR2.setText(ledger.getTitle());
                }
                RussiaList.add(russiaAR2);
            }
            if (title.contains("Перевод собственных средств. НДС не обл.")) {
                if (set.add(ledger.getNo())) {
                    ledger.setReasonCode("100");
                    ledger.setVendorCode(getAccount(ledger.getTaxCode(), banklist));
                    LedgerList.add(ledger);
                } else {
                    int f = LedgerList.indexOf(ledger);
                    LedgerList.remove(f);
                    ledger.setReasonCode("100");
                    ledger.setVendorCode(getAccount(ledger.getTaxCode(), banklist));
                    LedgerList.add(ledger);
                }
                russiaAR1.setNo(ledger.getNo());
                russiaAR1.setDocumentDate(ledger.getDate());
                russiaAR1.setType("SA");
                russiaAR1.setCompanyCode("62F0");
                russiaAR1.setPostingDate(ledger.getDate());
                int i = ledger.getDate().indexOf("/", 0);
                russiaAR1.setPeriod(ledger.getDate().substring(0, i));
                russiaAR1.setCurrencyRate("RUB");
                russiaAR1.setPostkey("50");
                russiaAR1.setGl("1002030710");
                russiaAR1.setAmount(ledger.getAmount());
                russiaAR1.setValueDate(ledger.getDate());
                if (ledger.getTitle().length() > 50) {
                    russiaAR1.setText(ledger.getTitle().substring(0, 50));
                } else {
                    russiaAR1.setText(ledger.getTitle());
                }
                russiaAR1.setReasonCode("100");
                RussiaList.add(russiaAR1);

                russiaAR2.setNo(ledger.getNo());
                russiaAR2.setDocumentDate(ledger.getDate());
                russiaAR2.setType("SA");
                russiaAR2.setCompanyCode("62F0");
                russiaAR2.setPostingDate(ledger.getDate());
                russiaAR2.setPeriod(ledger.getDate().substring(0, i));
                russiaAR2.setCurrencyRate("RUB");
                russiaAR2.setPostkey("40");
                //该处需要查询对照表
                russiaAR2.setGl(getAccount(ledger.getTaxCode(), banklist));
                russiaAR2.setAmount("*");
                russiaAR2.setValueDate(ledger.getDate());
                if (ledger.getTitle().length() > 50) {
                    russiaAR2.setText(ledger.getTitle().substring(0, 50));
                } else {
                    russiaAR2.setText(ledger.getTitle());
                }
                russiaAR2.setReasonCode("100");
                RussiaList.add(russiaAR2);
            }
            if (title.contains("Размещение депозита")) {
                if (set.add(ledger.getNo())) {
                    ledger.setReasonCode("100");
                    ledger.setVendorCode("1002030740");
                    LedgerList.add(ledger);
                } else {
                    int f = LedgerList.indexOf(ledger);
                    LedgerList.remove(f);
                    ledger.setReasonCode("100");
                    ledger.setVendorCode("1002030740");
                    LedgerList.add(ledger);
                }
                russiaAR1.setNo(ledger.getNo());
                russiaAR1.setDocumentDate(ledger.getDate());
                russiaAR1.setType("KZ");
                russiaAR1.setCompanyCode("62F0");
                russiaAR1.setPostingDate(ledger.getDate());
                int i = ledger.getDate().indexOf("/", 0);
                russiaAR1.setPeriod(ledger.getDate().substring(0, i));
                russiaAR1.setCurrencyRate("RUB");
                russiaAR1.setPostkey("50");
                russiaAR1.setGl("1002030710");
                russiaAR1.setAmount(ledger.getAmount());
                russiaAR1.setValueDate(ledger.getDate());
                if (ledger.getTitle().length() > 50) {
                    russiaAR1.setText(ledger.getTitle().substring(0, 50));
                } else {
                    russiaAR1.setText(ledger.getTitle());
                }
                russiaAR1.setReasonCode("100");
                RussiaList.add(russiaAR1);

                russiaAR2.setNo(ledger.getNo());
                russiaAR2.setDocumentDate(ledger.getDate());
                russiaAR2.setType("KZ");
                russiaAR2.setCompanyCode("62F0");
                russiaAR2.setPostingDate(ledger.getDate());
                russiaAR2.setPeriod(ledger.getDate().substring(0, i));
                russiaAR2.setCurrencyRate("RUB");
                russiaAR2.setPostkey("40");
                russiaAR2.setGl("1002030740");
                russiaAR2.setAmount("*");
                russiaAR2.setValueDate(ledger.getDate());
                if (ledger.getTitle().length() > 50) {
                    russiaAR2.setText(ledger.getTitle().substring(0, 50));
                } else {
                    russiaAR2.setText(ledger.getTitle());
                }
                russiaAR2.setReasonCode("100");
                RussiaList.add(russiaAR2);
            }
            if (title.contains("Покупка")) {
                if (set.add(ledger.getNo())) {
                    ledger.setReasonCode("100");
                    ledger.setVendorCode("1001010200");
                    LedgerList.add(ledger);
                } else {
                    int f = LedgerList.indexOf(ledger);
                    LedgerList.remove(f);
                    ledger.setReasonCode("100");
                    ledger.setVendorCode("1001010200");
                    LedgerList.add(ledger);
                }
                russiaAR1.setNo(ledger.getNo());
                russiaAR1.setDocumentDate(ledger.getDate());
                russiaAR1.setType("SA");
                russiaAR1.setCompanyCode("62F0");
                russiaAR1.setPostingDate(ledger.getDate());
                int i = ledger.getDate().indexOf("/", 0);
                russiaAR1.setPeriod(ledger.getDate().substring(0, i));
                russiaAR1.setCurrencyRate("RUB");
                russiaAR1.setPostkey("50");
                russiaAR1.setGl("1002030710");
                russiaAR1.setAmount(ledger.getAmount());
                russiaAR1.setValueDate(ledger.getDate());
                if (ledger.getTitle().length() > 50) {
                    russiaAR1.setText(ledger.getTitle().substring(0, 50));
                } else {
                    russiaAR1.setText(ledger.getTitle());
                }
                russiaAR1.setReasonCode("100");
                RussiaList.add(russiaAR1);

                russiaAR2.setNo(ledger.getNo());
                russiaAR2.setDocumentDate(ledger.getDate());
                russiaAR2.setType("SA");
                russiaAR2.setCompanyCode("62F0");
                russiaAR2.setPostingDate(ledger.getDate());
                russiaAR2.setPeriod(ledger.getDate().substring(0, i));
                russiaAR2.setCurrencyRate("RUB");
                russiaAR2.setPostkey("40");
                russiaAR2.setGl("1001010200");
                russiaAR2.setAmount("*");
                russiaAR2.setValueDate(ledger.getDate());
                if (ledger.getTitle().length() > 50) {
                    russiaAR2.setText(ledger.getTitle().substring(0, 50));
                } else {
                    russiaAR2.setText(ledger.getTitle());
                }
                russiaAR2.setReasonCode("100");
                RussiaList.add(russiaAR2);
            }
            if (title.contains("НДФЛ")) {
                if (set.add(ledger.getNo())) {
                    ledger.setReasonCode("164");
                    ledger.setVendorCode("V62B100502");
                    LedgerList.add(ledger);
                } else {
                    int f = LedgerList.indexOf(ledger);
                    LedgerList.remove(f);
                    ledger.setReasonCode("164");
                    ledger.setVendorCode("V62B100502");
                    LedgerList.add(ledger);
                }

                russiaAR1.setNo(ledger.getNo());
                russiaAR1.setDocumentDate(ledger.getDate());
                russiaAR1.setType("KZ");
                russiaAR1.setCompanyCode("62F0");
                russiaAR1.setPostingDate(ledger.getDate());
                int i = ledger.getDate().indexOf("/", 0);
                russiaAR1.setPeriod(ledger.getDate().substring(0, i));
                russiaAR1.setCurrencyRate("RUB");
                russiaAR1.setPostkey("50");
                russiaAR1.setGl("1002030710");
                russiaAR1.setAmount(ledger.getAmount());
                russiaAR1.setValueDate(ledger.getDate());
                if (ledger.getTitle().length() > 50) {
                    russiaAR1.setText(ledger.getTitle().substring(0, 50));
                } else {
                    russiaAR1.setText(ledger.getTitle());
                }
                russiaAR1.setReasonCode("164");
                RussiaList.add(russiaAR1);

                russiaAR2.setNo(ledger.getNo());
                russiaAR2.setDocumentDate(ledger.getDate());
                russiaAR2.setType("KZ");
                russiaAR2.setCompanyCode("62F0");
                russiaAR2.setPostingDate(ledger.getDate());
                russiaAR2.setPeriod(ledger.getDate().substring(0, i));
                russiaAR2.setCurrencyRate("RUB");
                russiaAR2.setSglInd("9");
                russiaAR2.setPostkey("29");
                russiaAR2.setGl("V62B100502");
                russiaAR2.setAmount("*");
                russiaAR2.setDueOn(ledger.getDate());
                if (ledger.getTitle().length() > 50) {
                    russiaAR2.setText(ledger.getTitle().substring(0, 50));
                } else {
                    russiaAR2.setText(ledger.getTitle());
                }
                russiaAR2.setReasonCode("164");
                RussiaList.add(russiaAR2);
            }
            if (title.contains("Платеж на страхование от НС")) {
                if (set.add(ledger.getNo())) {
                    ledger.setReasonCode("150");
                    ledger.setVendorCode("V62B100534");
                    LedgerList.add(ledger);
                } else {
                    int f = LedgerList.indexOf(ledger);
                    LedgerList.remove(f);
                    ledger.setReasonCode("150");
                    ledger.setVendorCode("V62B100534");
                    LedgerList.add(ledger);
                }

                russiaAR1.setNo(ledger.getNo());
                russiaAR1.setDocumentDate(ledger.getDate());
                russiaAR1.setType("KZ");
                russiaAR1.setCompanyCode("62F0");
                russiaAR1.setPostingDate(ledger.getDate());
                int i = ledger.getDate().indexOf("/", 0);
                russiaAR1.setPeriod(ledger.getDate().substring(0, i));
                russiaAR1.setCurrencyRate("RUB");
                russiaAR1.setPostkey("50");
                russiaAR1.setGl("1002030710");
                russiaAR1.setAmount(ledger.getAmount());
                russiaAR1.setValueDate(ledger.getDate());
                if (ledger.getTitle().length() > 50) {
                    russiaAR1.setText(ledger.getTitle().substring(0, 50));
                } else {
                    russiaAR1.setText(ledger.getTitle());
                }
                russiaAR1.setReasonCode("150");
                RussiaList.add(russiaAR1);

                russiaAR2.setNo(ledger.getNo());
                russiaAR2.setDocumentDate(ledger.getDate());
                russiaAR2.setType("KZ");
                russiaAR2.setCompanyCode("62F0");
                russiaAR2.setPostingDate(ledger.getDate());
                russiaAR2.setPeriod(ledger.getDate().substring(0, i));
                russiaAR2.setCurrencyRate("RUB");
                russiaAR2.setSglInd("3");
                russiaAR2.setPostkey("29");
                russiaAR2.setGl("V62B100534");
                russiaAR2.setAmount("*");
                russiaAR2.setDueOn(ledger.getDate());
                if (ledger.getTitle().length() > 50) {
                    russiaAR2.setText(ledger.getTitle().substring(0, 50));
                } else {
                    russiaAR2.setText(ledger.getTitle());
                }
                russiaAR2.setReasonCode("150");
                RussiaList.add(russiaAR2);
            }
            if (title.contains("Авансовые платежи зачисляемый ы ФСС")) {
                if (set.add(ledger.getNo())) {
                    ledger.setReasonCode("150");
                    ledger.setVendorCode("V62B100520");
                    LedgerList.add(ledger);
                } else {
                    int f = LedgerList.indexOf(ledger);
                    LedgerList.remove(f);
                    ledger.setReasonCode("150");
                    ledger.setVendorCode("V62B100520");
                    LedgerList.add(ledger);
                }

                russiaAR1.setNo(ledger.getNo());
                russiaAR1.setDocumentDate(ledger.getDate());
                russiaAR1.setType("KZ");
                russiaAR1.setCompanyCode("62F0");
                russiaAR1.setPostingDate(ledger.getDate());
                int i = ledger.getDate().indexOf("/", 0);
                russiaAR1.setPeriod(ledger.getDate().substring(0, i));
                russiaAR1.setCurrencyRate("RUB");
                russiaAR1.setPostkey("50");
                russiaAR1.setGl("1002030710");
                russiaAR1.setAmount(ledger.getAmount());
                russiaAR1.setValueDate(ledger.getDate());
                if (ledger.getTitle().length() > 50) {
                    russiaAR1.setText(ledger.getTitle().substring(0, 50));
                } else {
                    russiaAR1.setText(ledger.getTitle());
                }
                russiaAR1.setReasonCode("150");
                RussiaList.add(russiaAR1);

                russiaAR2.setNo(ledger.getNo());
                russiaAR2.setDocumentDate(ledger.getDate());
                russiaAR2.setType("KZ");
                russiaAR2.setCompanyCode("62F0");
                russiaAR2.setPostingDate(ledger.getDate());
                russiaAR2.setPeriod(ledger.getDate().substring(0, i));
                russiaAR2.setCurrencyRate("RUB");
                russiaAR2.setSglInd("3");
                russiaAR2.setPostkey("29");
                russiaAR2.setGl("V62B100520");
                russiaAR2.setAmount("*");
                russiaAR2.setDueOn(ledger.getDate());
                if (ledger.getTitle().length() > 50) {
                    russiaAR2.setText(ledger.getTitle().substring(0, 50));
                } else {
                    russiaAR2.setText(ledger.getTitle());
                }
                russiaAR2.setReasonCode("150");
                RussiaList.add(russiaAR2);
            }
            if (title.contains("Страховый взносы на ОМС зачисляемые в бюджет  ФФОМС")) {
                if (set.add(ledger.getNo())) {
                    ledger.setReasonCode("150");
                    ledger.setVendorCode("V62B100531");
                    LedgerList.add(ledger);
                } else {
                    int f = LedgerList.indexOf(ledger);
                    LedgerList.remove(f);
                    ledger.setReasonCode("150");
                    ledger.setVendorCode("V62B100531");
                    LedgerList.add(ledger);
                }

                russiaAR1.setNo(ledger.getNo());
                russiaAR1.setDocumentDate(ledger.getDate());
                russiaAR1.setType("KZ");
                russiaAR1.setCompanyCode("62F0");
                russiaAR1.setPostingDate(ledger.getDate());
                int i = ledger.getDate().indexOf("/", 0);
                russiaAR1.setPeriod(ledger.getDate().substring(0, i));
                russiaAR1.setCurrencyRate("RUB");
                russiaAR1.setPostkey("50");
                russiaAR1.setGl("1002030710");
                russiaAR1.setAmount(ledger.getAmount());
                russiaAR1.setValueDate(ledger.getDate());
                if (ledger.getTitle().length() > 50) {
                    russiaAR1.setText(ledger.getTitle().substring(0, 50));
                } else {
                    russiaAR1.setText(ledger.getTitle());
                }
                russiaAR1.setReasonCode("150");
                RussiaList.add(russiaAR1);

                russiaAR2.setNo(ledger.getNo());
                russiaAR2.setDocumentDate(ledger.getDate());
                russiaAR2.setType("KZ");
                russiaAR2.setCompanyCode("62F0");
                russiaAR2.setPostingDate(ledger.getDate());
                russiaAR2.setPeriod(ledger.getDate().substring(0, i));
                russiaAR2.setCurrencyRate("RUB");
                russiaAR2.setSglInd("3");
                russiaAR2.setPostkey("29");
                russiaAR2.setGl("V62B100531");
                russiaAR2.setAmount("*");
                russiaAR2.setDueOn(ledger.getDate());
                if (ledger.getTitle().length() > 50) {
                    russiaAR2.setText(ledger.getTitle().substring(0, 50));
                } else {
                    russiaAR2.setText(ledger.getTitle());
                }
                russiaAR2.setReasonCode("150");
                RussiaList.add(russiaAR2);
            }
            if (title.contains("Страх. Взносы на выплату строховой части трудовой пенси")) {
                if (set.add(ledger.getNo())) {
                    ledger.setReasonCode("150");
                    ledger.setVendorCode("V62B100523");
                    LedgerList.add(ledger);
                } else {
                    int f = LedgerList.indexOf(ledger);
                    LedgerList.remove(f);
                    ledger.setReasonCode("150");
                    ledger.setVendorCode("V62B100523");
                    LedgerList.add(ledger);
                }

                russiaAR1.setNo(ledger.getNo());
                russiaAR1.setDocumentDate(ledger.getDate());
                russiaAR1.setType("KZ");
                russiaAR1.setCompanyCode("62F0");
                russiaAR1.setPostingDate(ledger.getDate());
                int i = ledger.getDate().indexOf("/", 0);
                russiaAR1.setPeriod(ledger.getDate().substring(0, i));
                russiaAR1.setCurrencyRate("RUB");
                russiaAR1.setPostkey("50");
                russiaAR1.setGl("1002030710");
                russiaAR1.setAmount(ledger.getAmount());
                russiaAR1.setValueDate(ledger.getDate());
                if (ledger.getTitle().length() > 50) {
                    russiaAR1.setText(ledger.getTitle().substring(0, 50));
                } else {
                    russiaAR1.setText(ledger.getTitle());
                }
                russiaAR1.setReasonCode("150");
                RussiaList.add(russiaAR1);

                russiaAR2.setNo(ledger.getNo());
                russiaAR2.setDocumentDate(ledger.getDate());
                russiaAR2.setType("KZ");
                russiaAR2.setCompanyCode("62F0");
                russiaAR2.setPostingDate(ledger.getDate());
                russiaAR2.setPeriod(ledger.getDate().substring(0, i));
                russiaAR2.setCurrencyRate("RUB");
                russiaAR2.setSglInd("3");
                russiaAR2.setPostkey("29");
                russiaAR2.setGl("V62B100523");
                russiaAR2.setAmount("*");
                russiaAR2.setDueOn(ledger.getDate());
                if (ledger.getTitle().length() > 50) {
                    russiaAR2.setText(ledger.getTitle().substring(0, 50));
                } else {
                    russiaAR2.setText(ledger.getTitle());
                }
                russiaAR2.setReasonCode("150");
                RussiaList.add(russiaAR2);
            }
//            Vendor依据customer内容通过匹配表匹配
            else {
//                台账
                if (set.add(ledger.getNo())) {
                    if (getOtherAccount(ledger.getVendor(),otherBankList).size()>0)
                    {
                        ledger.setReasonCode(getOtherAccount(ledger.getVendor(),otherBankList).get(1).toString());
                        ledger.setVendorCode(getOtherAccount(ledger.getVendor(),otherBankList).get(0).toString());
                    }else {
                        ledger.setReasonCode("");
                        ledger.setVendorCode("");
                    }
                    LedgerList.add(ledger);
                } else {
                    int f = LedgerList.indexOf(ledger);
                    LedgerList.remove(f);
                    if (getOtherAccount(ledger.getVendor(),otherBankList).size()>0)
                    {
                        ledger.setReasonCode(getOtherAccount(ledger.getVendor(),otherBankList).get(1).toString());
                        ledger.setVendorCode(getOtherAccount(ledger.getVendor(),otherBankList).get(0).toString());
                    }else {
                        ledger.setReasonCode("");
                        ledger.setVendorCode("");
                    }
                    LedgerList.add(ledger);
                }
//                FB01
                russiaAR1.setNo(ledger.getNo());
                russiaAR1.setDocumentDate(ledger.getDate());
                russiaAR1.setType("KZ");
                russiaAR1.setCompanyCode("62F0");
                russiaAR1.setPostingDate(ledger.getDate());
                int i = ledger.getDate().indexOf("/", 0);
                russiaAR1.setPeriod(ledger.getDate().substring(0, i));
                russiaAR1.setCurrencyRate("RUB");
                russiaAR1.setPostkey("50");
                russiaAR1.setGl("1002030710");
                russiaAR1.setAmount(ledger.getAmount());
                russiaAR1.setValueDate(ledger.getDate());
                if (ledger.getTitle().length() > 50) {
                    russiaAR1.setText(cleanTitle(ledger.getTitle().substring(0, 50)).substring(1));
                } else {
                    russiaAR1.setText(cleanTitle(ledger.getTitle()).substring(1));
                }
                if (getOtherAccount(ledger.getVendor(),otherBankList).size()>0)
                {
                    russiaAR1.setReasonCode(getOtherAccount(ledger.getVendor(),otherBankList).get(1).toString());
                }else {
                    russiaAR1.setReasonCode("");
                }
                if (!(russiaAR1.getGl().equals("")) && !(russiaAR1.getReasonCode().equals(""))){
                    RussiaList.add(russiaAR1);
                }


                russiaAR2.setNo(ledger.getNo());
                russiaAR2.setDocumentDate(ledger.getDate());
                russiaAR2.setType("KZ");
                russiaAR2.setCompanyCode("62F0");
                russiaAR2.setPostingDate(ledger.getDate());
                russiaAR2.setPeriod(ledger.getDate().substring(0, i));
                russiaAR2.setCurrencyRate("RUB");
                russiaAR2.setSglInd("3");
                russiaAR2.setPostkey("25");
                if (getOtherAccount(ledger.getVendor(),otherBankList).size()>0)
                {
                    russiaAR2.setGl(getOtherAccount(ledger.getVendor(),otherBankList).get(0).toString());
                }else {
                    russiaAR2.setGl("");
                }
                russiaAR2.setAmount("*");
                russiaAR2.setDueOn(ledger.getDate());

                if (ledger.getTitle().length() > 50) {
                    russiaAR2.setText(cleanTitle(ledger.getTitle()).substring(0, 50).substring(1));
                } else {
                    russiaAR2.setText(cleanTitle(ledger.getTitle()).substring(1));
                }

                russiaAR2.setReasonCode("");

                if (!(russiaAR2.getGl().equals(""))){
                    RussiaList.add(russiaAR2);
                }

            }
//

            if (set.add(ledger.getNo())) {
                LedgerList.add(ledger);
            }
        }
        System.out.println("101付款处理之后的台账：" + LedgerList.size());
        excelOutput_ledger(LedgerList, ledger_Path);
        return RussiaList;
    }

    private static String cleanTitle(String title) {
        if (title.length()>50){
            title = title.replaceAll("Оплата по счету","Оп. сч");
            title = title.replaceAll("счет","сч");
            title = title.substring(1);
        }
        return title.trim();
    }

    /**
     * 102 收款FB01 数据生成
     *
     * @param LedList
     * @return
     * @throws Exception
     */
    public static List HandleData2(List<RussiaLedger> LedList, String ledger_Path) throws Exception {
        List<RussiaAR> RussiaList = new ArrayList<>();
        for (RussiaLedger ledger : LedList) {
            ledger.setGl("1002030270");
            ledger.setSapClearNo("无需清账");
            RussiaAR russiaAR1 = new RussiaAR();
            RussiaAR russiaAR2 = new RussiaAR();
            String title = ledger.getTitle().replace(" ", "");
            if (title.contains("Выплата процентов по депозиту")) {
                russiaAR1.setNo(ledger.getNo());
                russiaAR1.setDocumentDate(ledger.getDate());
                russiaAR1.setType("DZ");
                russiaAR1.setCompanyCode("62F0");
                russiaAR1.setPostingDate(ledger.getDate());
                int i = ledger.getDate().indexOf("/", 0);
                russiaAR1.setPeriod(ledger.getDate().substring(0, i));
                russiaAR1.setCurrencyRate("USD");
                russiaAR1.setPostkey("40");
                russiaAR1.setGl("1002030270");
                russiaAR1.setAmount(ledger.getAmount());
                russiaAR1.setValueDate(ledger.getDate());
                russiaAR1.setText("Выплата процентов по депозиту");
                russiaAR1.setReasonCode("332");
                RussiaList.add(russiaAR1);

                russiaAR2.setNo(ledger.getNo());
                russiaAR2.setDocumentDate(ledger.getDate());
                russiaAR2.setType("DZ");
                russiaAR2.setCompanyCode("62F0");
                russiaAR2.setPostingDate(ledger.getDate());
                russiaAR2.setPeriod(ledger.getDate().substring(0, i));
                russiaAR2.setCurrencyRate("USD");
                russiaAR2.setPostkey("50");
                russiaAR2.setGl("5503010000");
                russiaAR2.setAmount("*");
                russiaAR2.setCostCenter("662F120000");
                russiaAR2.setText("Выплата процентов по депозиту");
                RussiaList.add(russiaAR2);

                ledger.setVendorCode("5503010000");
                ledger.setReasonCode("332");
                ledger.setText("Выплата процентов по депозиту");

            }
            if (title.contains("Возврат депозита по заявке")) {
                russiaAR1.setNo(ledger.getNo());
                russiaAR1.setDocumentDate(ledger.getDate());
                russiaAR1.setType("DZ");
                russiaAR1.setCompanyCode("62F0");
                russiaAR1.setPostingDate(ledger.getDate());
                int i = ledger.getDate().indexOf("/", 0);
                russiaAR1.setPeriod(ledger.getDate().substring(0, i));
                russiaAR1.setCurrencyRate("USD");
                russiaAR1.setPostkey("40");
                russiaAR1.setGl("1002030270");
                russiaAR1.setAmount(ledger.getAmount());
                russiaAR1.setValueDate(ledger.getDate());
                russiaAR1.setText("Возврат депозита по заявке");
                russiaAR1.setReasonCode("100");
                RussiaList.add(russiaAR1);

                russiaAR2.setNo(ledger.getNo());
                russiaAR2.setDocumentDate(ledger.getDate());
                russiaAR2.setType("DZ");
                russiaAR2.setCompanyCode("62F0");
                russiaAR2.setPostingDate(ledger.getDate());
                russiaAR2.setPeriod(ledger.getDate().substring(0, i));
                russiaAR2.setCurrencyRate("USD");
                russiaAR2.setPostkey("50");
                russiaAR2.setGl("1002030283");
                russiaAR2.setAmount("*");
                russiaAR2.setValueDate(ledger.getDate());
                russiaAR2.setText("Возврат депозита по заявке");
                russiaAR2.setReasonCode("100");
                RussiaList.add(russiaAR2);

                ledger.setVendorCode("1002030283");
                ledger.setReasonCode("100");
                ledger.setText("Возврат депозита по заявке");
            }
            if (title.contains("Покупка")) {
                int k = title.indexOf("Покупка");
                int j = title.indexOf("RUR");
                String str = title.substring(k, j + 3);

//                String uuid = UUID.randomUUID().toString().replaceAll("-","");
                russiaAR1.setNo(ledger.getNo());
                russiaAR1.setDocumentDate(ledger.getDate());
                russiaAR1.setType("DZ");
                russiaAR1.setCompanyCode("62F0");
                russiaAR1.setPostingDate(ledger.getDate());
                int i = ledger.getDate().indexOf("/", 0);
                russiaAR1.setPeriod(ledger.getDate().substring(0, i));
                russiaAR1.setCurrencyRate("USD");
                russiaAR1.setPostkey("40");
                russiaAR1.setGl("1002030270");
                russiaAR1.setAmount(ledger.getAmount());
                russiaAR1.setValueDate(ledger.getDate());
                russiaAR1.setText(str);
                russiaAR1.setReasonCode("100");
                RussiaList.add(russiaAR1);

                russiaAR2.setNo(ledger.getNo());
                russiaAR2.setDocumentDate(ledger.getDate());
                russiaAR2.setType("DZ");
                russiaAR2.setCompanyCode("62F0");
                russiaAR2.setPostingDate(ledger.getDate());
                russiaAR2.setPeriod(ledger.getDate().substring(0, i));
                russiaAR2.setCurrencyRate("USD");
                russiaAR2.setPostkey("50");
                russiaAR2.setGl("1001010200");
                russiaAR2.setAmount("*");
                russiaAR2.setValueDate(ledger.getDate());
                russiaAR2.setText(str);
                russiaAR2.setReasonCode("100");
                RussiaList.add(russiaAR2);

                ledger.setVendorCode("1001010200");
                ledger.setReasonCode("100");
                ledger.setText(str);
            }
            if (title.contains("перенос денежных средств от")) {
                int k = title.indexOf("*");
                int j = title.lastIndexOf("*");
                String str = title.substring(k + 1, j);
                russiaAR1.setNo(ledger.getNo());
                russiaAR1.setDocumentDate(ledger.getDate());
                russiaAR1.setType("DZ");
                russiaAR1.setCompanyCode("62F0");
                russiaAR1.setPostingDate(ledger.getDate());
                int i = ledger.getDate().indexOf("/", 0);
                russiaAR1.setPeriod(ledger.getDate().substring(0, i));
                russiaAR1.setCurrencyRate("USD");
                russiaAR1.setPostkey("40");
                russiaAR1.setGl("1002030270");
                russiaAR1.setAmount(ledger.getAmount());
                russiaAR1.setValueDate(ledger.getDate());
                russiaAR1.setText("Возврат депозита по заявке" + " " + str);
                russiaAR1.setReasonCode("100");
                RussiaList.add(russiaAR1);

                russiaAR2.setNo(ledger.getNo());
                russiaAR2.setDocumentDate(ledger.getDate());
                russiaAR2.setType("DZ");
                russiaAR2.setCompanyCode("62F0");
                russiaAR2.setPostingDate(ledger.getDate());
                russiaAR2.setPeriod(ledger.getDate().substring(0, i));
                russiaAR2.setCurrencyRate("USD");
                russiaAR2.setPostkey("50");
                russiaAR2.setGl("1002030280");
                russiaAR2.setAmount("*");
                russiaAR2.setValueDate(ledger.getDate());
                russiaAR2.setText("Возврат депозита по заявке" + " " + str);
                russiaAR2.setReasonCode("100");
                RussiaList.add(russiaAR2);

                ledger.setVendorCode("1002030280");
                ledger.setReasonCode("100");
                ledger.setText("Возврат депозита по заявке" + " " + str);
            }
        }
        excelOutput_ledger2(LedList, ledger_Path);
        return RussiaList;

    }

    /**
     * 102 付款 FB01数据生成
     *
     * @param LedList
     * @param ledger_Path
     * @return
     * @throws Exception
     */
    public static List HandleData02(List<RussiaLedger> LedList, String ledger_Path, String filePath) throws Exception {
        List<RussiaAR> RussiaList = new ArrayList<>();
        for (RussiaLedger ledger : LedList) {
            ledger.setGl("1002030270");
            ledger.setSapClearNo("无需清账");
            RussiaAR russiaAR1 = new RussiaAR();
            RussiaAR russiaAR2 = new RussiaAR();
            String title = ledger.getTitle().replace(" ", "");
            if (title.contains("Комиссия")) {
                russiaAR1.setNo(ledger.getNo());
                russiaAR1.setDocumentDate(ledger.getDate());
                russiaAR1.setType("KZ");
                russiaAR1.setCompanyCode("62F0");
                russiaAR1.setPostingDate(ledger.getDate());
                int i = ledger.getDate().indexOf("/", 0);
                russiaAR1.setPeriod(ledger.getDate().substring(0, i));
                russiaAR1.setCurrencyRate("USD");
                russiaAR1.setPostkey("50");
                russiaAR1.setGl("1002030270");
                russiaAR1.setAmount(ledger.getAmount());
                russiaAR1.setValueDate(ledger.getDate());
                russiaAR1.setText("оплата банк комиссии");
                russiaAR1.setReasonCode("362");
                RussiaList.add(russiaAR1);

                russiaAR2.setNo(ledger.getNo());
                russiaAR2.setDocumentDate(ledger.getDate());
                russiaAR2.setType("KZ");
                russiaAR2.setCompanyCode("62F0");
                russiaAR2.setPostingDate(ledger.getDate());
                russiaAR2.setPeriod(ledger.getDate().substring(0, i));
                russiaAR2.setCurrencyRate("USD");
                russiaAR2.setPostkey("40");
                russiaAR2.setGl("5503040000");
                russiaAR2.setAmount("*");
                russiaAR2.setCostCenter("662F120000");
                russiaAR2.setText("оплата банк комиссии");
                RussiaList.add(russiaAR2);

                ledger.setVendorCode("5503040000");
                ledger.setReasonCode("362");
                ledger.setText("оплата банк комиссии");
            }
            if (title.contains("Labour contract")) {
                russiaAR1.setNo(ledger.getNo());
                russiaAR1.setDocumentDate(ledger.getDate());
                russiaAR1.setType("KZ");
                russiaAR1.setCompanyCode("62F0");
                russiaAR1.setPostingDate(ledger.getDate());
                int i = ledger.getDate().indexOf("/", 0);
                russiaAR1.setPeriod(ledger.getDate().substring(0, i));
                russiaAR1.setCurrencyRate("USD");
                russiaAR1.setPostkey("50");
                russiaAR1.setGl("1002030270");
                russiaAR1.setAmount(ledger.getAmount());
                russiaAR1.setValueDate(ledger.getDate());
                russiaAR1.setText("выплата аванса зп");
                russiaAR1.setReasonCode("150");
                RussiaList.add(russiaAR1);

                russiaAR2.setNo(ledger.getNo());
                russiaAR2.setDocumentDate(ledger.getDate());
                russiaAR2.setType("KZ");
                russiaAR2.setCompanyCode("62F0");
                russiaAR2.setPostingDate(ledger.getDate());
                russiaAR2.setPeriod(ledger.getDate().substring(0, i));
                russiaAR2.setCurrencyRate("USD");
                russiaAR2.setPostkey("40");
                russiaAR2.setGl("2151010000");
                russiaAR2.setAmount("*");
                russiaAR2.setText("выплата аванса зп");
                RussiaList.add(russiaAR2);

                ledger.setVendorCode("2151010000");
                ledger.setReasonCode("150");
                ledger.setText("выплата аванса зп");
            }
            if (title.contains("Размещение депозита")) {
                russiaAR1.setNo(ledger.getNo());
                russiaAR1.setDocumentDate(ledger.getDate());
                russiaAR1.setType("KZ");
                russiaAR1.setCompanyCode("62F0");
                russiaAR1.setPostingDate(ledger.getDate());
                int i = ledger.getDate().indexOf("/", 0);
                russiaAR1.setPeriod(ledger.getDate().substring(0, i));
                russiaAR1.setCurrencyRate("USD");
                russiaAR1.setPostkey("50");
                russiaAR1.setGl("1002030270");
                russiaAR1.setAmount(ledger.getAmount());
                russiaAR1.setValueDate(ledger.getDate());

                russiaAR1.setText("Размещение депозита");

                russiaAR1.setReasonCode("100");
                RussiaList.add(russiaAR1);

                russiaAR2.setNo(ledger.getNo());
                russiaAR2.setDocumentDate(ledger.getDate());
                russiaAR2.setType("KZ");
                russiaAR2.setCompanyCode("62F0");
                russiaAR2.setPostingDate(ledger.getDate());
                russiaAR2.setPeriod(ledger.getDate().substring(0, i));
                russiaAR2.setCurrencyRate("USD");
                russiaAR2.setPostkey("40");
                russiaAR2.setGl("1002030283");
                russiaAR2.setAmount("*");
                russiaAR2.setValueDate(ledger.getDate());
                russiaAR2.setText("Размещение депозита");
                russiaAR2.setReasonCode("100");
                RussiaList.add(russiaAR2);

                ledger.setVendorCode("1002030283");
                ledger.setReasonCode("100");
                ledger.setText("Размещение депозита");
            }
            if ("".equals(ledger.getVendorCode()) || ledger.getVendorCode() == null) {
                List<CustomerCode> vendorList = getCustomer_Vendor(filePath);
                List<CustomerCode> cusCode = new ArrayList<>();
                String vendorCode = "";
                for (CustomerCode customer : vendorList) {
                    if (ledger.getVendor().contains(customer.getName())) {
                        vendorCode = customer.getCustomerCode();
                        cusCode.add(customer);
                    }
                }
                if (cusCode.size() == 1) {
                    russiaAR1.setNo(ledger.getNo());
                    russiaAR1.setDocumentDate(ledger.getDate());
                    russiaAR1.setType("KZ");
                    russiaAR1.setCompanyCode("62F0");
                    russiaAR1.setPostingDate(ledger.getDate());
                    int i = ledger.getDate().indexOf("/", 0);
                    russiaAR1.setPeriod(ledger.getDate().substring(0, i));
                    russiaAR1.setCurrencyRate("USD");
                    russiaAR1.setPostkey("50");
                    russiaAR1.setGl("1002030270");
                    russiaAR1.setAmount(ledger.getAmount());
                    russiaAR1.setValueDate(ledger.getDate());
                    russiaAR1.setReasonCode("160");
                    RussiaList.add(russiaAR1);

                    russiaAR2.setNo(ledger.getNo());
                    russiaAR2.setDocumentDate(ledger.getDate());
                    russiaAR2.setType("KZ");
                    russiaAR2.setCompanyCode("62F0");
                    russiaAR2.setPostingDate(ledger.getDate());
                    russiaAR2.setPeriod(ledger.getDate().substring(0, i));
                    russiaAR2.setCurrencyRate("USD");
                    russiaAR2.setPostkey("25");
                    russiaAR2.setGl(vendorCode);
                    russiaAR2.setAmount("*");
                    RussiaList.add(russiaAR2);

                    ledger.setVendorCode(vendorCode);
                    ledger.setReasonCode("160");
                }
            }
        }
        excelOutput_ledger(LedList, ledger_Path);
        return RussiaList;
    }

    /**
     * 103 收款数据生成FB01
     *
     * @param LedList
     * @param ledger_Path
     * @param filePath
     * @return
     * @throws Exception
     */
    public static List HandleData3(List<RussiaLedger> LedList, String ledger_Path, String filePath) throws Exception {
        List<RussiaAR> RussiaList = new ArrayList<>();
        for (RussiaLedger ledger : LedList) {
            ledger.setGl("1002030280");
            ledger.setSapClearNo("无需清账");
            RussiaAR russiaAR1 = new RussiaAR();
            RussiaAR russiaAR2 = new RussiaAR();
            String title = ledger.getTitle().replace(" ", "");
            //获取对照表数据
            List<Contrast> contrastList = getContrast(filePath);
            for (Contrast contrast : contrastList) {
                if (ledger.getTitle().contains(contrast.getText())) {
//                    System.out.println("输出title:" + title);
                    russiaAR1.setNo(ledger.getNo());
                    russiaAR1.setDocumentDate(ledger.getDate());
                    russiaAR1.setType("DZ");
                    russiaAR1.setCompanyCode("62F0");
                    russiaAR1.setPostingDate(ledger.getDate());
                    int i = ledger.getDate().indexOf("/", 0);
                    russiaAR1.setPeriod(ledger.getDate().substring(0, i));
                    russiaAR1.setCurrencyRate("USD");
                    russiaAR1.setPostkey("40");
                    russiaAR1.setGl("1002030280");
                    russiaAR1.setAmount(ledger.getAmount());
                    russiaAR1.setValueDate(ledger.getDate());
                    if ("поступление от Арена S KZT".equals(contrast.getText())) {
                        int k = title.indexOf("*");
                        int j = title.lastIndexOf("*");
                        russiaAR1.setText(title.substring(k + 1, j));
                    } else {
                        russiaAR1.setText(contrast.getText());
                    }
                    russiaAR1.setReasonCode(contrast.getReasonCode());
                    RussiaList.add(russiaAR1);

                    russiaAR2.setNo(ledger.getNo());
                    russiaAR2.setDocumentDate(ledger.getDate());
                    russiaAR2.setType("DZ");
                    russiaAR2.setCompanyCode("62F0");
                    russiaAR2.setPostingDate(ledger.getDate());
                    russiaAR2.setPeriod(ledger.getDate().substring(0, i));
                    russiaAR2.setCurrencyRate("USD");
                    if ("customer".equals(contrast.getType())) {
                        russiaAR2.setPostkey("15");
                    }
                    if ("vendor".equals(contrast.getType())) {
                        russiaAR2.setPostkey("35");
                    }
                    russiaAR2.setGl(contrast.getCustomerCode());
                    russiaAR2.setAmount("*");
                    russiaAR2.setPmtBlock("B");
                    if ("поступление от Арена S KZT".equals(contrast.getText())) {
                        int k = title.indexOf("*");
                        int j = title.lastIndexOf("*");
                        russiaAR2.setText(title.substring(k + 1, j));
                        ledger.setText(title.substring(k + 1, j));
                    } else {
                        russiaAR2.setText(contrast.getText());
                        ledger.setText(contrast.getText());
                    }
                    RussiaList.add(russiaAR2);

                    ledger.setVendorCode(contrast.getCustomerCode());
                    ledger.setReasonCode(contrast.getReasonCode());
                }
            }
        }
        excelOutput_ledger2(LedList, ledger_Path);
        return RussiaList;
    }

    /**
     * 106 编号数据 付款数据处理
     *
     * @param LedList
     * @param ledger_Path
     * @return
     */
    public static List HandleData6(List<RussiaLedger> LedList, String ledger_Path) throws Exception {
        List<RussiaAR> RussiaList = new ArrayList<>();
        for (RussiaLedger ledger : LedList) {
            ledger.setGl("1002030720");
            ledger.setSapClearNo("无需清账");
            RussiaAR russiaAR1 = new RussiaAR();
            RussiaAR russiaAR2 = new RussiaAR();
            String title = ledger.getTitle().replace(" ", "");
            if (title.contains("ДЕ 2010")) {
                russiaAR1.setNo(ledger.getNo());
                russiaAR1.setDocumentDate(ledger.getDate());
                russiaAR1.setType("KZ");
                russiaAR1.setCompanyCode("62F0");
                russiaAR1.setPostingDate(ledger.getDate());
                int i = ledger.getDate().indexOf("/", 0);
                russiaAR1.setPeriod(ledger.getDate().substring(0, i));
                russiaAR1.setCurrencyRate("RUB");
                russiaAR1.setPostkey("50");
                russiaAR1.setGl("1002030720");
                russiaAR1.setAmount(ledger.getAmount());
                russiaAR1.setValueDate(ledger.getDate());
                russiaAR1.setText("Уплата пошлины 2010");
                russiaAR1.setReasonCode("162");
                RussiaList.add(russiaAR1);

                russiaAR2.setNo(ledger.getNo());
                russiaAR2.setDocumentDate(ledger.getDate());
                russiaAR2.setType("KZ");
                russiaAR2.setCompanyCode("62F0");
                russiaAR2.setPostingDate(ledger.getDate());
                russiaAR2.setPeriod(ledger.getDate().substring(0, i));
                russiaAR2.setCurrencyRate("RUB");
                russiaAR2.setPostkey("25");
                russiaAR2.setGl("V800000585");
                russiaAR2.setAmount("*");
                russiaAR2.setText("Уплата пошлины 2010");
                RussiaList.add(russiaAR2);

                ledger.setVendorCode("V800000585");
                ledger.setReasonCode("162");
                ledger.setText("Уплата пошлины 2010");
            }
            if (title.contains("ДЕ 1010")) {
                russiaAR1.setNo(ledger.getNo());
                russiaAR1.setDocumentDate(ledger.getDate());
                russiaAR1.setType("KZ");
                russiaAR1.setCompanyCode("62F0");
                russiaAR1.setPostingDate(ledger.getDate());
                int i = ledger.getDate().indexOf("/", 0);
                russiaAR1.setPeriod(ledger.getDate().substring(0, i));
                russiaAR1.setCurrencyRate("RUB");
                russiaAR1.setPostkey("50");
                russiaAR1.setGl("1002030720");
                russiaAR1.setAmount(ledger.getAmount());
                russiaAR1.setValueDate(ledger.getDate());
                russiaAR1.setText("Уплата сборов 1010");
                russiaAR1.setReasonCode("162");
                RussiaList.add(russiaAR1);

                russiaAR2.setNo(ledger.getNo());
                russiaAR2.setDocumentDate(ledger.getDate());
                russiaAR2.setType("KZ");
                russiaAR2.setCompanyCode("62F0");
                russiaAR2.setPostingDate(ledger.getDate());
                russiaAR2.setPeriod(ledger.getDate().substring(0, i));
                russiaAR2.setCurrencyRate("RUB");
                russiaAR2.setPostkey("25");
                russiaAR2.setGl("V800000585");
                russiaAR2.setAmount("*");
                russiaAR2.setText("Уплата сборов 1010");
                RussiaList.add(russiaAR2);

                ledger.setVendorCode("V800000585");
                ledger.setReasonCode("162");
                ledger.setText("Уплата сборов 1010");
            }
            if (title.contains("ДЕ 5010")) {
                russiaAR1.setNo(ledger.getNo());
                russiaAR1.setDocumentDate(ledger.getDate());
                russiaAR1.setType("KZ");
                russiaAR1.setCompanyCode("62F0");
                russiaAR1.setPostingDate(ledger.getDate());
                int i = ledger.getDate().indexOf("/", 0);
                russiaAR1.setPeriod(ledger.getDate().substring(0, i));
                russiaAR1.setCurrencyRate("RUB");
                russiaAR1.setPostkey("50");
                russiaAR1.setGl("1002030720");
                russiaAR1.setAmount(ledger.getAmount());
                russiaAR1.setValueDate(ledger.getDate());
                russiaAR1.setText("Уплата аванса 5010");
                russiaAR1.setReasonCode("160");
                RussiaList.add(russiaAR1);

                russiaAR2.setNo(ledger.getNo());
                russiaAR2.setDocumentDate(ledger.getDate());
                russiaAR2.setType("KZ");
                russiaAR2.setCompanyCode("62F0");
                russiaAR2.setPostingDate(ledger.getDate());
                russiaAR2.setPeriod(ledger.getDate().substring(0, i));
                russiaAR2.setCurrencyRate("RUB");
                russiaAR2.setPostkey("25");
                russiaAR2.setGl("V800000586");
                russiaAR2.setAmount("*");
                russiaAR2.setText("Уплата аванса 5010");
                RussiaList.add(russiaAR2);

                ledger.setVendorCode("V800000586");
                ledger.setReasonCode("160");
                ledger.setText("Уплата аванса 5010");
            }
        }
        excelOutput_ledger(LedList, ledger_Path);
        return RussiaList;
    }

    /**
     * 获取对照表的数据
     * 103 收款类型数据的对照表
     *
     * @param filePath1
     * @return
     * @throws Exception
     */
    public static List getContrast(String filePath1) throws Exception {
        File excelFile = new File(filePath1);
        FileInputStream EIL_file_IO = new FileInputStream(excelFile);
        Workbook wb = WorkbookFactory.create(EIL_file_IO);
        Sheet sheet = wb.getSheetAt(0);
        List<Contrast> contrastList = new ArrayList<>();
        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            String customerCode = "";
            String reasonCode = "";
            String type = "";
            String text = "";
            Row row = sheet.getRow(r);
            if (row.getCell(0) != null) {
                row.getCell(0).setCellType(Cell.CELL_TYPE_STRING);
                customerCode = row.getCell(0).getStringCellValue();
            }
            if (row.getCell(4) != null) {
                row.getCell(4).setCellType(Cell.CELL_TYPE_STRING);
                reasonCode = row.getCell(4).getStringCellValue();
            }
            if (row.getCell(5) != null) {
                row.getCell(5).setCellType(Cell.CELL_TYPE_STRING);
                type = row.getCell(5).getStringCellValue();
            }
            if (row.getCell(6) != null) {
                row.getCell(6).setCellType(Cell.CELL_TYPE_STRING);
                text = row.getCell(6).getStringCellValue();
            }
            Contrast contrast = new Contrast();
            contrast.setCustomerCode(customerCode);
            contrast.setType(type);
            contrast.setReasonCode(reasonCode);
            contrast.setText(text);
            if (!"".equals(customerCode)) {
                contrastList.add(contrast);
            }
        }
        return contrastList;
    }

    /**
     * 获取对照表
     * 102 扣款类型数据的对照表
     *
     * @param filePath1
     * @return
     * @throws Exception
     */
    public static List getCustomer_Vendor(String filePath1) throws Exception {
        List<CustomerCode> vendorList = new ArrayList<>();
        File excelFile = new File(filePath1);
        FileInputStream EIL_file_IO = new FileInputStream(excelFile);
        Workbook wb = WorkbookFactory.create(EIL_file_IO);
        Sheet sheet = wb.getSheetAt(0);
        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            String name = "";
            String vendorCode = "";
            Row row = sheet.getRow(r);
            if (row.getCell(0) != null) {
                row.getCell(0).setCellType(Cell.CELL_TYPE_STRING);
                vendorCode = row.getCell(0).getStringCellValue();
            }
            if (row.getCell(1) != null) {
                row.getCell(1).setCellType(Cell.CELL_TYPE_STRING);
                name = row.getCell(1).getStringCellValue();
            }
            CustomerCode customerCode = new CustomerCode();
            customerCode.setName(name);
            customerCode.setCustomerCode(vendorCode);
            if (!"".equals(name)) {
                vendorList.add(customerCode);
            }
        }
        return vendorList;
    }

    /**
     * 获取银行水单数据 -- 付款
     *
     * @param filePath
     */
    public static List getLedgerFunction(String filePath) throws Exception {
        System.out.println("数据获取中！");
        File excelFile = new File(filePath);
        FileInputStream EIL_file_IO = new FileInputStream(excelFile);
        Workbook wb = WorkbookFactory.create(EIL_file_IO);
        Sheet sheet = wb.getSheetAt(0);
        List<RussiaLedger> LedList = new ArrayList<>();
        for (int r = 10; r <= sheet.getLastRowNum(); r++) {
            String date = "";
            String vendor = "";
            String title = "";
            String taxCode = "";
            String amount = "";
            Row row = sheet.getRow(r);
            RussiaLedger ledger = new RussiaLedger();
            if (row.getCell(11) != null) {
                row.getCell(11).setCellType(Cell.CELL_TYPE_STRING);
                String am = clearData(row.getCell(11).getStringCellValue());
                if (!"".equals(am)) {
                    BigDecimal bd = new BigDecimal(am);
                    amount = bd.setScale(2, BigDecimal.ROUND_HALF_UP).toPlainString();
                }
            }
            if (!"".equals(amount) && Double.parseDouble(amount) != 0) {//判断amount是否为0
                amount = row.getCell(11).getStringCellValue();
                if (row.getCell(0) != null) {
                    row.getCell(0).setCellType(Cell.CELL_TYPE_STRING);
                    date = clearData(row.getCell(0).getStringCellValue());
                    date = date.replace(".", "/");
                    date = date.substring(3, 6) + date.substring(0, 3) + date.substring(6);
                }
                if (row.getCell(2) != null) {
                    row.getCell(2).setCellType(Cell.CELL_TYPE_STRING);
                    title = row.getCell(2).getStringCellValue();
                }
                if (row.getCell(5) != null) {
                    row.getCell(5).setCellType(Cell.CELL_TYPE_STRING);
                    vendor = row.getCell(5).getStringCellValue();
                }
                if (row.getCell(9) != null) {
                    row.getCell(9).setCellType(Cell.CELL_TYPE_STRING);
                    taxCode = clearData(row.getCell(9).getStringCellValue());
                }
                String uuid = UUID.randomUUID().toString().replaceAll("-", "");
                ledger.setNo(uuid);
                ledger.setAmount(amount.replace(" ", ""));
                ledger.setDate(date);
                ledger.setTitle(title);
                ledger.setVendor(vendor);
                ledger.setTaxCode(taxCode);
                if (!"".equals(title)) {
                    LedList.add(ledger);
                }
            }
        }
        return LedList;
    }

    /**
     * 获取银行水单数据 -- 收款
     *
     * @param filePath
     * @return
     * @throws Exception
     */
    public static List getLedgerFunction2(String filePath) throws Exception {
        System.out.println("数据获取中！");
        File excelFile = new File(filePath);
        FileInputStream EIL_file_IO = new FileInputStream(excelFile);
        Workbook wb = WorkbookFactory.create(EIL_file_IO);
        Sheet sheet = wb.getSheetAt(0);
        List<RussiaLedger> LedList = new ArrayList<>();
        for (int r = 10; r <= sheet.getLastRowNum(); r++) {
            String date = "";
            String vendor = "";
            String title = "";
            String taxCode = "";
            String amount = "";
            Row row = sheet.getRow(r);
            RussiaLedger ledger = new RussiaLedger();
            if (row.getCell(12) != null) {
                row.getCell(12).setCellType(Cell.CELL_TYPE_STRING);
                String am = clearData(row.getCell(12).getStringCellValue());
                if (!"".equals(am)) {
                    BigDecimal bd = new BigDecimal(am);
                    amount = bd.setScale(2, BigDecimal.ROUND_HALF_UP).toPlainString();
                }
            }
            if (!"".equals(amount) && Double.parseDouble(amount) != 0) {//判断amount是否为0
                amount = row.getCell(12).getStringCellValue();
                if (row.getCell(0) != null) {
                    row.getCell(0).setCellType(Cell.CELL_TYPE_STRING);
                    date = clearData(row.getCell(0).getStringCellValue());
                    date = date.replace(".", "/");
                    date = date.substring(3, 6) + date.substring(0, 3) + date.substring(6);
                }
                if (row.getCell(2) != null) {
                    row.getCell(2).setCellType(Cell.CELL_TYPE_STRING);
                    title = row.getCell(2).getStringCellValue();
                }
                if (row.getCell(5) != null) {
                    row.getCell(5).setCellType(Cell.CELL_TYPE_STRING);
                    vendor = row.getCell(5).getStringCellValue();
                }
                if (row.getCell(9) != null) {
                    row.getCell(9).setCellType(Cell.CELL_TYPE_STRING);
                    taxCode = clearData(row.getCell(9).getStringCellValue());
                }
                String uuid = UUID.randomUUID().toString().replaceAll("-", "");
                ledger.setNo(uuid);
                ledger.setAmount(amount.replace(" ", ""));
                ledger.setDate(date);
                ledger.setTitle(title);
                ledger.setVendor(vendor);
                ledger.setTaxCode(taxCode);
                if (!"".equals(title)) {
                    LedList.add(ledger);
                }
            }
        }
        return LedList;
    }

    /**
     * 提取有效数字
     *
     * @param data
     * @return
     */
    public static String clearData(String data) {
        if (!"".equals(data)) {
            Pattern p = Pattern.compile("[^.0-9]");//提取有效数字
            data = p.matcher(data).replaceAll("").trim();
        }
        return data;
    }

    /**
     * 生成FB01表格
     *
     * @param RussiaARList
     * @param filePath_FB
     * @throws Exception
     */
    public static void excelOutput_FB01(List<RussiaAR> RussiaARList, String filePath_FB) throws Exception {
        //创建表格
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        //定义第一个sheet页
        XSSFSheet xssfSheet = xssfWorkbook.createSheet("Sheet1");
        //第一个sheet页数据（生成台账表）
        Row row0 = xssfSheet.createRow(0);
        String[] headers = new String[]{"No", "Document date", "Type", "Company Code", "Posting Date", "Period", "Currency/Rate", "Document Number", "Translatn Date", "Reference", "Cross-CC No.", "Doc.Header Text", "Branch Number", "Trading Part.BA", "Postkey", "GL", "SGL Ind", "Amount", "Business Area", "Cost Center", "Profit Center", "Value Date", "Due On", "Pmt Block", "Tax Code", "Calculate Tax", "Issue Date", "Ext. No", "Bank/Acct No", "Disc Base", "Assignment", "Text", "Reason code"};
        for (int i = 0; i < headers.length; i++) {
            xssfSheet.setColumnWidth(i, 5000);
            XSSFCell cell = (XSSFCell) row0.createCell(i);
            XSSFRichTextString text = new XSSFRichTextString(headers[i]);
            cell.setCellValue(text);
        }
        int rowNum = 1;
        for (RussiaAR russiaAR : RussiaARList) {
            XSSFRow row = xssfSheet.createRow(rowNum);
            xssfSheet.setColumnWidth(rowNum, 5000);
            row.createCell(0).setCellValue(russiaAR.getNo());
            row.createCell(1).setCellValue(russiaAR.getDocumentDate());
            row.createCell(2).setCellValue(russiaAR.getType());
            row.createCell(3).setCellValue(russiaAR.getCompanyCode());
            row.createCell(4).setCellValue(russiaAR.getPostingDate());
            row.createCell(5).setCellValue(russiaAR.getPeriod());
            row.createCell(6).setCellValue(russiaAR.getCurrencyRate());
            row.createCell(7).setCellValue(russiaAR.getDocumentNumber());
            row.createCell(8).setCellValue(russiaAR.getTranslatnDate());

            row.createCell(9).setCellValue(russiaAR.getReference());
            row.createCell(10).setCellValue(russiaAR.getCrossCCNo());
            row.createCell(11).setCellValue(russiaAR.getDocHeaderText());
            row.createCell(12).setCellValue(russiaAR.getBrancNumber());
            row.createCell(13).setCellValue(russiaAR.getTradingPartBA());
            row.createCell(14).setCellValue(russiaAR.getPostkey());
            row.createCell(15).setCellValue(russiaAR.getGl());
            row.createCell(16).setCellValue(russiaAR.getSglInd());
            row.createCell(17).setCellValue(russiaAR.getAmount());

            row.createCell(18).setCellValue(russiaAR.getBusinessArea());
            row.createCell(19).setCellValue(russiaAR.getCostCenter());
            row.createCell(20).setCellValue(russiaAR.getProfitCenter());
            row.createCell(21).setCellValue(russiaAR.getValueDate());
            row.createCell(22).setCellValue(russiaAR.getDueOn());
            row.createCell(23).setCellValue(russiaAR.getPmtBlock());
            row.createCell(24).setCellValue(russiaAR.getTaxCode());
            row.createCell(25).setCellValue(russiaAR.getCalculateTax());
            row.createCell(26).setCellValue(russiaAR.getIssueDate());

            row.createCell(27).setCellValue(russiaAR.getExtNo());
            row.createCell(28).setCellValue(russiaAR.getBankAcctNo());
            row.createCell(29).setCellValue(russiaAR.getDiscBase());
            row.createCell(30).setCellValue(russiaAR.getAssignment());
            row.createCell(31).setCellValue(russiaAR.getText());
            row.createCell(32).setCellValue(russiaAR.getReasonCode());
            rowNum++;
        }
        FileOutputStream fileOutputStream = new FileOutputStream(filePath_FB);
        xssfWorkbook.write(fileOutputStream);
    }

    /**
     * 生成台账--付款
     *
     * @param LedgerList
     * @param ledger_Path
     */
    public static void excelOutput_ledger(List<RussiaLedger> LedgerList, String ledger_Path) throws Exception {
        System.out.println("生成付款台账！");
        //创建表格
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        //定义第一个sheet页
        XSSFSheet xssfSheet = xssfWorkbook.createSheet("Sheet1");
        //第一个sheet页数据（生成台账表）
        Row row0 = xssfSheet.createRow(0);
        String[] headers = new String[]{"No", "Date", "Vendor", "Title", "Code", "Amount", "Reason Code", "SAPIncomeNo.", "SpecialG/L", "Invoice", "SAPClearNo.", "Status", "GL"};
        for (int i = 0; i < headers.length; i++) {
            xssfSheet.setColumnWidth(i, 5000);
            XSSFCell cell = (XSSFCell) row0.createCell(i);
            XSSFRichTextString text = new XSSFRichTextString(headers[i]);
            cell.setCellValue(text);
        }
        int rowNum = 1;
        for (RussiaLedger ledger : LedgerList) {
            XSSFRow row = xssfSheet.createRow(rowNum);
            xssfSheet.setColumnWidth(rowNum, 5000);
            row.createCell(0).setCellValue(ledger.getNo());
            row.createCell(1).setCellValue(ledger.getDate());
            row.createCell(2).setCellValue(ledger.getVendor());
            row.createCell(3).setCellValue(ledger.getTitle());
            row.createCell(4).setCellValue(ledger.getVendorCode());
            row.createCell(5).setCellValue(ledger.getAmount());
            row.createCell(6).setCellValue(ledger.getReasonCode());
            row.createCell(7).setCellValue(ledger.getSapIncomeNo());
            row.createCell(8).setCellValue(ledger.getSpecialGL());
            row.createCell(9).setCellValue(ledger.getInvoice());
            row.createCell(10).setCellValue(ledger.getSapClearNo());
            row.createCell(11).setCellValue(ledger.getStatus());
            row.createCell(12).setCellValue(ledger.getGl());
            rowNum++;
        }
        FileOutputStream fileOutputStream = new FileOutputStream(ledger_Path);
        xssfWorkbook.write(fileOutputStream);
    }

    /**
     * 生成台账--收款
     *
     * @param LedgerList
     * @param ledger_Path
     */
    public static void excelOutput_ledger2(List<RussiaLedger> LedgerList, String ledger_Path) throws Exception {
        System.out.println("生成收款台账！");
        //创建表格
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        //定义第一个sheet页
        XSSFSheet xssfSheet = xssfWorkbook.createSheet("Sheet1");
        //第一个sheet页数据（生成台账表）
        Row row0 = xssfSheet.createRow(0);
        String[] headers = new String[]{"NO", "Date", "Customer", "Title", "Code", "Amount", "Reason Code", "Text", "Invoice", "SAPIncomeNo.", "SAPClearNo.", "Status", "ReceivedInfor.Date", "KeyinDate(ClearDate)", "Comment", "GL"};
        for (int i = 0; i < headers.length; i++) {
            xssfSheet.setColumnWidth(i, 5000);
            XSSFCell cell = (XSSFCell) row0.createCell(i);
            XSSFRichTextString text = new XSSFRichTextString(headers[i]);
            cell.setCellValue(text);
        }
        int rowNum = 1;
        for (RussiaLedger ledger : LedgerList) {
            XSSFRow row = xssfSheet.createRow(rowNum);
            xssfSheet.setColumnWidth(rowNum, 5000);
            row.createCell(0).setCellValue(ledger.getNo());
            row.createCell(1).setCellValue(ledger.getDate());
            row.createCell(2).setCellValue(ledger.getVendor());
            row.createCell(3).setCellValue(ledger.getTitle());
            row.createCell(4).setCellValue(ledger.getVendorCode());
            row.createCell(5).setCellValue(ledger.getAmount());
            row.createCell(6).setCellValue(ledger.getReasonCode());
            row.createCell(7).setCellValue(ledger.getText());
            row.createCell(8).setCellValue(ledger.getInvoice());
            row.createCell(9).setCellValue(ledger.getSapIncomeNo());
            row.createCell(10).setCellValue(ledger.getSapClearNo());
            row.createCell(11).setCellValue(ledger.getStatus());
            row.createCell(12).setCellValue("");
            row.createCell(13).setCellValue("");
            row.createCell(14).setCellValue("");
            row.createCell(15).setCellValue(ledger.getGl());
            rowNum++;
        }
        FileOutputStream fileOutputStream = new FileOutputStream(ledger_Path);
        xssfWorkbook.write(fileOutputStream);
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
                String[] vl = title.split("( |ф)");
                List<String> list1 = Arrays.asList(vl);
                List<String> list2 = new ArrayList<>();
                for (String s : list1){
                    Pattern pattern = Pattern.compile(".*\\d+.*");
                    Matcher m = pattern.matcher(s);
                    if (m.matches()) {//判断是否有数字
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
}
