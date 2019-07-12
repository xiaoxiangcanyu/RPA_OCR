package HttpUtil;

import DataClean.BaseUtil;
import DataClean.MiddleEastAll;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Pattern;

/**
 *银行流水单表生成台账表
 */
public class MiddleEastBS_AR extends BaseUtil{
    public static void main(String[] args) {
        String filePath = args[0];//原始银行水单数据表
        String filePath_ledger = args[1];//生成台账表路径
        String filePath2 = args[2];//生成余额表路径
//        String filePath = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\中东AR\\test\\20190305\\20190401\\HSBC-AED.xlsx";//原始银行水单数据表
//        String filePath_ledger = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\中东AR\\test\\20190305\\20190401\\AR_Matching_List.xlsx";//生成台账表路径
//        String filePath2 = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\中东AR\\test\\20190305\\20190401\\balance_excel.xls";//生成余额表路径
        //截取银行类型
        String bankCode = filePath.substring(filePath.lastIndexOf("\\") + 1, filePath.lastIndexOf("."));
        System.out.println("银行类型:" + bankCode);
        try {
            List<MiddleEastAll> MiddleEastALLList = null;
            switch (bankCode) {
                case "BOC-USD":
                case "BOC-AED":
                    MiddleEastALLList = getMiddleEastExcel_BOC(filePath, bankCode);
                    break;
                case "HFC-USD":
                    MiddleEastALLList = getMiddleEastExcel_HFC(filePath, bankCode);
                    break;
                case "HSBC-AED":
                case "HSBC-USD":
                    MiddleEastALLList = getMiddleEastExcel_HSBC(filePath, bankCode);
                    break;
                //SCB银行存在多币种混合的情况，最后不匹配的银行即为SCB银行水单
                default:
                    MiddleEastALLList = getMiddleEastExcel_SCB(filePath, bankCode);
                    break;
            }
            if (MiddleEastALLList.size() > 0){
                //判断台账表格是否存在
                File file = new File(filePath_ledger);
                if(!file.exists()) {
                    int index = 1;
                    for (MiddleEastAll middleEastAll : MiddleEastALLList){
                        middleEastAll.setId(String.valueOf(index));
                        index++;
                    }
                    //若不存在，创建新的台账
                    excelOutput_Log(MiddleEastALLList, filePath_ledger);
                }else {
                    //存在，则取出台账中原有数据与新数据合并
                    List<MiddleEastAll>  oldMiddleEastList = getLedgerExcelData(filePath_ledger);
                    System.out.println(oldMiddleEastList.size());
                    for (MiddleEastAll middleEastAll:oldMiddleEastList){
                        System.out.println("id:"+middleEastAll.getId());
                    }

                    //排序
                    oldMiddleEastList = cleanSortData(oldMiddleEastList);
                    String num = oldMiddleEastList.get(oldMiddleEastList.size()-1).getId();
                    int i = Integer.parseInt(num) + 1;
                    for (MiddleEastAll middleEastAll : MiddleEastALLList){
                        middleEastAll.setId(String.valueOf(i));
                        i++;
                    }
                    oldMiddleEastList.addAll(MiddleEastALLList);
                    excelOutput_Log(oldMiddleEastList, filePath_ledger);
                }
                //获取银行水单余额
                //获取各个币种的水单余额
                Set set = new HashSet();//定义set集合保存银行币种
                Map<String, List<MiddleEastAll>> map = new HashMap<>();//存储不同币种对应的所有数据
                for (MiddleEastAll middleEastAll : MiddleEastALLList){
                    if (set.add(middleEastAll.getCurrency())){
                        List<MiddleEastAll> list = new ArrayList<>();
                        for (MiddleEastAll middleEastAll2 : MiddleEastALLList){
                            if (middleEastAll.getCurrency().equals(middleEastAll2.getCurrency())){
                                list.add(middleEastAll2);
                            }
                        }
                        map.put(middleEastAll.getCurrency(),list);
                    }
                }
                List<MiddleEastAll> MEbalance = new ArrayList<>();
                for (String key : map.keySet()){
                    //调用方法获取余额数据
                    MEbalance.addAll(getBalance(map.get(key)));
                }
                //生成余额表
                excelOutput_Balance(MEbalance,filePath2);
                System.out.println("运行结束！");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 获取银行水单余额
     * @param list
     * @return
     * @throws Exception
     */
    public static List getBalance(List<MiddleEastAll> list)throws Exception{
        //获取银行水单的最后余额
        List<MiddleEastAll> MEbalance = new ArrayList<>();
        SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
        String balance = "";
        String bank = "";
        if (list.size() == 1){
            MiddleEastAll middleEastAll = new MiddleEastAll();
            middleEastAll.setBalance(convertData(list.get(0).getBalance()));
            middleEastAll.setBank(list.get(0).getBank());
            middleEastAll.setDocumentDate(sdf.format(new Date()));
            MEbalance.add(middleEastAll);
            System.out.println(list.get(0).getBank() + "--" + sdf.format(new Date()) + "--" + convertData(list.get(0).getBalance()));
        }else {
            for (int i = 1; i < list.size(); i++){
                bank = list.get(0).getBank();
                Date date1 = sdf.parse(list.get(0).getDocumentDate());
                Date date2 = sdf.parse(list.get(i).getDocumentDate());
                if (date1.getTime() >= date2.getTime()) {
                    balance = list.get(0).getBalance();
                }else {
                    balance = list.get(i).getBalance();
                }
            }
            MiddleEastAll middleEastAll = new MiddleEastAll();
            middleEastAll.setBalance(convertData(balance));
            middleEastAll.setBank(bank);
            middleEastAll.setDocumentDate(sdf.format(new Date()));
            MEbalance.add(middleEastAll);
            System.out.println(bank + "--" + sdf.format(new Date()) + "--" + balance);
        }
        return MEbalance;
    }

    public static List getMiddleEastExcel_BOC(String filePath, String bankCode) throws Exception {
        File excelFile = new File(filePath);
        String documentDate = "";
        String income = "";//正向
        String charge = "";//负向
        String summary = "";
        String type = "";
        String remark = "";
        String balance = "";
        FileInputStream EIL_file_IO = new FileInputStream(excelFile);
        Workbook wb = WorkbookFactory.create(EIL_file_IO);
        Sheet sheet = wb.getSheetAt(0);
        List<MiddleEastAll> MiddleEastList = new ArrayList<>();
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
        SimpleDateFormat sdf1 = new SimpleDateFormat("MM/dd/yyyy");
        for (int r = 4; r <= sheet.getLastRowNum(); r++) {
            Row rows = sheet.getRow(r);//获取第r行
            if (rows.getCell(2) != null) {
                rows.getCell(2).setCellType(Cell.CELL_TYPE_STRING);
                String money = rows.getCell(2).getStringCellValue();
                if (money.contains("-")) {
                    charge = clearData(rows.getCell(2).getStringCellValue());
                    income = "";
                } else {
                    income = clearData(rows.getCell(2).getStringCellValue());
                    charge = "";
                }
            } else {
                income = "";
                charge = "";
            }
            if (rows.getCell(3) != null) {
                rows.getCell(3).setCellType(Cell.CELL_TYPE_STRING);
                balance = clearData(rows.getCell(3).getStringCellValue());
            } else {
                balance = "";
            }
            if (rows.getCell(5) != null) {
                rows.getCell(5).setCellType(Cell.CELL_TYPE_STRING);
                documentDate = rows.getCell(5).getStringCellValue();
                documentDate = sdf1.format(sdf.parse(documentDate));
            } else {
                documentDate = "";
            }
            if (rows.getCell(8) != null) {
                rows.getCell(8).setCellType(Cell.CELL_TYPE_STRING);
                type = rows.getCell(8).getStringCellValue();
            } else {
                type = "";
            }
            if (rows.getCell(9) != null) {
                rows.getCell(9).setCellType(Cell.CELL_TYPE_STRING);
                remark = rows.getCell(9).getStringCellValue();
            } else {
                remark = "";
            }
            MiddleEastAll middleEastAll = new MiddleEastAll();
            if ("BOC-AED".equals(bankCode) ) {
                middleEastAll.setBank("BOC-AED 1002011701");
                middleEastAll.setCurrency("AED");
            }else if ("BOC-USD".equals(bankCode)){
                middleEastAll.setCurrency("USD");
                middleEastAll.setBank("BOC-USD 1002012001");
            }
            middleEastAll.setCharge(charge);
            middleEastAll.setIncome(income);
            middleEastAll.setDocumentDate(documentDate);
            middleEastAll.setBalance(balance);
            summary = remark + " " + type;
            middleEastAll.setSummary(summary.trim());
            MiddleEastList.add(middleEastAll);
        }
        List<MiddleEastAll> MiddleEastALLList = new ArrayList<>();
        for (MiddleEastAll middleEastAll : MiddleEastList) {
            if (middleEastAll.getDocumentDate() != null && !"".equals(middleEastAll.getDocumentDate())) {
                //判断数据类型
                middleEastAll = ckeckTtAndLc(middleEastAll, bankCode);
                if (checkData(middleEastAll.getSummary())){
                    MiddleEastALLList.add(middleEastAll);
                }
            }
        }
        return MiddleEastALLList;
    }

    public static List getMiddleEastExcel_HFC(String filePath,String bankCode) throws Exception {
        File excelFile = new File(filePath);
        String documentDate = "";
        String income = "";//正向
        String charge = "";//负向
        String remark = "";
        String balance = "";
        FileInputStream EIL_file_IO = new FileInputStream(excelFile);
        Workbook wb = WorkbookFactory.create(EIL_file_IO);
        Sheet sheet = wb.getSheetAt(0);
        List<MiddleEastAll> MiddleEastList = new ArrayList<>();
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
        SimpleDateFormat sdf1 = new SimpleDateFormat("MM/dd/yyyy");
        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row rows = sheet.getRow(r);//获取第r行
            if (rows.getCell(0) != null){
                rows.getCell(0).setCellType(Cell.CELL_TYPE_STRING);
                documentDate = rows.getCell(0).getStringCellValue();
                documentDate = documentDate.replace("-","/");
                documentDate = sdf1.format(sdf.parse(documentDate));
            }else {
                documentDate = "";
            }
            if (rows.getCell(2) != null){
                rows.getCell(2).setCellType(Cell.CELL_TYPE_STRING);
                charge = clearData(rows.getCell(2).getStringCellValue());
            }else {
                charge = "";
            }
            if (rows.getCell(3) != null){
                rows.getCell(3).setCellType(Cell.CELL_TYPE_STRING);
                income = clearData(rows.getCell(3).getStringCellValue());
            }else {
                income = "";
            }
            if (rows.getCell(4) != null){
                rows.getCell(4).setCellType(Cell.CELL_TYPE_STRING);
                balance = clearData(rows.getCell(4).getStringCellValue());
            }else {
                balance = "";
            }
            if (rows.getCell(5) != null){
                rows.getCell(5).setCellType(Cell.CELL_TYPE_STRING);
                remark = rows.getCell(5).getStringCellValue();
            }else {
                remark = "";
            }
            MiddleEastAll middleEastAll = new MiddleEastAll();
            middleEastAll.setBank("HFC-USD 1002200252");
            middleEastAll.setCurrency("USD");
            middleEastAll.setIncome(income);
            middleEastAll.setBalance(balance);
            middleEastAll.setCharge(charge);
            middleEastAll.setSummary(remark);
            middleEastAll.setDocumentDate(documentDate);
            MiddleEastList.add(middleEastAll);
        }
        List<MiddleEastAll> MiddleEastALLList = new ArrayList<>();
        for (MiddleEastAll middleEastAll : MiddleEastList){
            //判断数据类型
            if (middleEastAll.getDocumentDate() != null && !"".equals(middleEastAll.getDocumentDate())) {
                middleEastAll = ckeckTtAndLc(middleEastAll,bankCode);
                if (checkData(middleEastAll.getSummary())){
                    MiddleEastALLList.add(middleEastAll);
                }
            }
        }
        return MiddleEastALLList;
    }

    public static List getMiddleEastExcel_HSBC(String filePath,String bankCode) throws Exception {
        String documentDate = "";
        String income = "";//正向
        String charge = "";//负向
        String summary = "";
        String currency = "";
        String balance = "";
        File excelFile = new File(filePath);
        FileInputStream EIL_file_IO = new FileInputStream(excelFile);
        Workbook wb = WorkbookFactory.create(EIL_file_IO);
        Sheet sheet = wb.getSheetAt(0);
        List<MiddleEastAll> MiddleEastList = new ArrayList<>();
        SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
        SimpleDateFormat sdf1 = new SimpleDateFormat("MM/dd/yyyy");
        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row rows = sheet.getRow(r);//获取第r行
            if (rows.getCell(3) != null){
                rows.getCell(3).setCellType(Cell.CELL_TYPE_STRING);
                currency = rows.getCell(3).getStringCellValue();
            }else {
                currency = "";
            }
            if (rows.getCell(10) != null){
                rows.getCell(10).setCellType(Cell.CELL_TYPE_STRING);
                summary = rows.getCell(10).getStringCellValue();
            }else {
                summary = "";
            }
            if (rows.getCell(13) != null){
                rows.getCell(13).setCellType(Cell.CELL_TYPE_STRING);
                documentDate = rows.getCell(13).getStringCellValue();
                documentDate = sdf1.format(sdf.parse(documentDate));
            } else {
                documentDate = "";
            }
            if (rows.getCell(14) != null){
                rows.getCell(14).setCellType(Cell.CELL_TYPE_STRING);
                income = clearData(rows.getCell(14).getStringCellValue());
            }else {
                income = "";
            }
            if (rows.getCell(15) != null){
                rows.getCell(15).setCellType(Cell.CELL_TYPE_STRING);
                charge = clearData(rows.getCell(15).getStringCellValue());
            }else {
                charge = "";
            }
            if (rows.getCell(16) != null){
                rows.getCell(16).setCellType(Cell.CELL_TYPE_STRING);
                balance = clearData(rows.getCell(16).getStringCellValue());
            }else {
                balance = "";
            }
            MiddleEastAll middleEastAll = new MiddleEastAll();
//            String uuid = UUID.randomUUID().toString().replaceAll("-","");
//            middleEastAll.setId(uuid);
            middleEastAll.setCurrency(currency);
            middleEastAll.setCharge(charge);
            middleEastAll.setIncome(income);
            middleEastAll.setBalance(balance);
            middleEastAll.setDocumentDate(documentDate);
            middleEastAll.setSummary(summary);
            if ("AED".equals(currency)){
                middleEastAll.setBank("HSBC-AED 1002031711");
            }else if ("USD".equals(currency)){
                middleEastAll.setBank("HSBC-USD 1002030262");
            }
            MiddleEastList.add(middleEastAll);
        }
        List<MiddleEastAll> MiddleEastALLList = new ArrayList<>();
        for (MiddleEastAll middleEastAll : MiddleEastList){
            //判断数据类型
            if (middleEastAll.getDocumentDate() != null && !"".equals(middleEastAll.getDocumentDate())) {
                middleEastAll = ckeckTtAndLc(middleEastAll,bankCode);
                if (checkData(middleEastAll.getSummary())){
                    MiddleEastALLList.add(middleEastAll);
                }
            }
        }
        return MiddleEastALLList;
    }
    public static List getMiddleEastExcel_SCB(String filePath,String bankCode) throws Exception {
        String documentDate = "";
        String income = "";//正向
        String charge = "";//负向
        String summary = "";
        String currency = "";
        String state = "";
        String balance = "";
        File excelFile = new File(filePath);
        FileInputStream EIL_file_IO = new FileInputStream(excelFile);
        Workbook wb = WorkbookFactory.create(EIL_file_IO);
        Sheet sheet = wb.getSheetAt(0);
        List<MiddleEastAll> MiddleEastList = new ArrayList<>();
        SimpleDateFormat sdf1 = new SimpleDateFormat("MM/dd/yyyy");
        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row rows = sheet.getRow(r);//获取第r行
            if (rows.getCell(2) != null){
                rows.getCell(2).setCellType(Cell.CELL_TYPE_STRING);
                currency = rows.getCell(2).getStringCellValue();
            }else {
                currency = "";
            }
            if (rows.getCell(5) != null){
                rows.getCell(5).setCellType(Cell.CELL_TYPE_STRING);
                balance = clearData(rows.getCell(5).getStringCellValue());
            }else {
                balance = "";
            }
            if (rows.getCell(9) != null){
                if (rows.getCell(9).getCellType() == 0){
                    if (DateUtil.isCellDateFormatted(rows.getCell(9))){//判断是否是时间类型
                        documentDate = sdf1.format(rows.getCell(9).getDateCellValue());
                    }
                }
            } else {
                documentDate = "";
            }
            if (rows.getCell(11) != null){
                rows.getCell(11).setCellType(Cell.CELL_TYPE_STRING);
                state = rows.getCell(11).getStringCellValue();
            }else {
                state ="";
            }
            if (state != null && !"".equals(state)){
                if ("C".equals(state)){
                    if (rows.getCell(10) != null){
                        rows.getCell(10).setCellType(Cell.CELL_TYPE_STRING);
                        income = clearData(rows.getCell(10).getStringCellValue());
                        charge = "";
                    }else {
                        income = "";
                        charge = "";
                    }
                }else if ("D".equals(state)){
                    if (rows.getCell(10) != null){
                        rows.getCell(10).setCellType(Cell.CELL_TYPE_STRING);
                        charge = clearData(rows.getCell(10).getStringCellValue());
                        income = "";
                    }else {
                        charge = "";
                        income = "";
                    }
                }
            }
            if (rows.getCell(16) != null){
                rows.getCell(16).setCellType(Cell.CELL_TYPE_STRING);
                summary = rows.getCell(16).getStringCellValue();
            } else {
                summary = "";
            }
            MiddleEastAll middleEastAll = new MiddleEastAll();
//            String uuid = UUID.randomUUID().toString().replaceAll("-","");
//            middleEastAll.setId(uuid);
            middleEastAll.setCurrency(currency);
            middleEastAll.setCharge(charge);
            middleEastAll.setIncome(income);
            middleEastAll.setBalance(balance);
            middleEastAll.setDocumentDate(documentDate);
            middleEastAll.setSummary(summary);
            if ("AED".equals(currency)){
                middleEastAll.setBank("SCB-AED 1002071720");
            }else if ("USD".equals(currency)){
                middleEastAll.setBank("SCB-USD 1002070231");
            }else if ("EUR".equals(currency)){
                middleEastAll.setBank("SCB-EUR 1002070320");
            }
            MiddleEastList.add(middleEastAll);
        }
        List<MiddleEastAll> MiddleEastALLList = new ArrayList<>();
        for (MiddleEastAll middleEastAll : MiddleEastList){
            //判断数据类型
            if (middleEastAll.getDocumentDate() != null && !"".equals(middleEastAll.getDocumentDate())) {
                middleEastAll = ckeckTtAndLc(middleEastAll,bankCode);
                if (checkData(middleEastAll.getSummary())){
                    MiddleEastALLList.add(middleEastAll);
                }
            }
        }
        return MiddleEastALLList;
    }

    /**
     * 判断数据是否为一下数据类型
     * @param summary
     * @return
     */
    public static boolean checkData(String summary){
        boolean flag = false;
        if (!summary.contains("汇出汇款") && !summary.contains("国际汇出汇款")
                && summary.indexOf("SCN",0) != 0){
            flag = true;
        }
        return flag;
    }

    /**
     * 数据处理
     * @param data
     * @return
     */
    public static String clearData(String data) {
        if (!"".equals(data)){
            Pattern p = Pattern.compile("[^./0-9]");//提取有效数字
            data = p.matcher(data).replaceAll("").trim();
        }
        return data;
    }

    /**
     * 数据排序
     * @param MiddleEastList
     * @return
     */
    public static List<MiddleEastAll> cleanSortData( List<MiddleEastAll> MiddleEastList) {
        if (MiddleEastList.size() > 0){
            Collections.sort(MiddleEastList, (p1, p2) -> {
                System.out.println("p1:"+p1.getId()+","+p1.getDocumentDate());
                System.out.println("p2:"+p2.getId()+","+p2.getDocumentDate());
                if(Integer.parseInt(p1.getId()) > Integer.parseInt(p2.getId())){
                    return 1;
                }
                if(Integer.parseInt(p1.getId()) == Integer.parseInt(p2.getId())){
                    return 0;
                }
                return -1;
            });
        }
        return MiddleEastList;
    }
    /**
     * 判断数据类型数TT/LC
     * @return
     */
    public static MiddleEastAll ckeckTtAndLc(MiddleEastAll middleEastAll, String bankCode){
        String summary = middleEastAll.getSummary();
        switch (bankCode){
            case "BOC-AED":
            case "BOC-USD":
                if (!"".equals(summary) && summary != null){
                    if (summary.indexOf("BP") == 0 || summary.indexOf("AD") == 0 ||summary.indexOf("GA") == 0 || summary.indexOf("收费") == 0 || summary.indexOf("收税") == 0
                    || summary.indexOf("CB") == 0 || summary.indexOf("NEGO CHG VAT") == 0 || summary.indexOf("TIME DEPOSIT") == 0 || summary.indexOf("VAT OF") == 0
                    || summary.indexOf("TI") == 0 || summary.contains("IW TT COMMISSIONS VAT") || summary.contains("LOCAL COURIER CHARGE")
                    || summary.contains("BANK CONFIRMATION CHARGE") || summary.contains("转账支出")){
                        middleEastAll.setTtLCMark("LC");
                    }
                    if (summary.contains("国际汇入汇款") || summary.indexOf("IW FTS") == 0 || summary.indexOf("OC") == 0) {
                        middleEastAll.setTtLCMark("TT");
                    }
                    if (summary.contains("汇入汇款解付")){
                        middleEastAll.setTtLCMark("TT/LC");
                    }
                }
                break;
            case "HFC-USD":
                if (!"".equals(summary) && summary != null){
                    if (summary.contains("手续费") && summary.contains("GEMS产生的手续费")){
                        middleEastAll.setTtLCMark("LC");
                    }
                    if (summary.contains("汇入汇款解付")){
                        middleEastAll.setTtLCMark("TT/LC");
                    }
                }
                break;
            case "HSBC-AED":
            case "HSBC-USD":
                if (!"".equals(summary) && summary != null){
                    if (summary.indexOf("DC ADVISING COMM") == 0 || summary.indexOf("PAY BY") == 0 || summary.indexOf("EXP BILL") == 0 || summary.indexOf("BPC") == 0
                            || summary.indexOf("BAC") == 0 || summary.indexOf("5 VAT-TXN DTD") == 0 || summary.indexOf("FNGJBA") == 0 || summary.indexOf("SEE BELOW") == 0){
                        middleEastAll.setTtLCMark("LC");
                    }
                    if (summary.contains("NONREF CDM-CC")){
                        middleEastAll.setTtLCMark("TT");
                    }
                    if (summary.contains("汇入汇款解付")){
                        middleEastAll.setTtLCMark("TT/LC");
                    }
                }
                break;
            default:
                if (!"".equals(summary) && summary != null){
                    if (summary.indexOf("BILL NO") == 0 || summary.indexOf("123") == 0 || summary.indexOf("EZJ") == 0 || summary.indexOf("WPSCHGS  Post Transaction") == 0
                            || summary.indexOf("SAE") == 0 || summary.indexOf("AUDIT CONFIRMATION FEE") == 0 || summary.indexOf("VALUE ADDED TAX") == 0
                            || summary.indexOf("VAT  Post transaction narration manually") == 0 ){
                        middleEastAll.setTtLCMark("LC");
                    }
                    if (summary.indexOf("CDM CASH") == 0 || summary.indexOf("LOCAL CHEQUE DEPOSIT") == 0 || summary.indexOf("CASH DEPOSIT") == 0 || summary.indexOf("IL") == 0
                            || summary.indexOf("IT") == 0 || summary.indexOf("VAT CHRG") == 0 || summary.indexOf("VAT ON CHARGES  IT") == 0){
                        middleEastAll.setTtLCMark("TT");
                    }
                    if (summary.contains("汇入汇款解付")){
                        middleEastAll.setTtLCMark("TT/LC");
                    }
                }
                break;
        }
        return middleEastAll;
    }

    /**
     * 生成余额表
     * @param MiddleEastALLList
     * @param filePath
     * @throws Exception
     */
    public static void excelOutput_Balance(List<MiddleEastAll> MiddleEastALLList, String filePath) throws Exception{
        //创建表格
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        //定义第一个sheet页
        XSSFSheet xssfSheet = xssfWorkbook.createSheet("Sheet1");
        //第一个sheet页数据（生成台账表）
        Row row0 = xssfSheet.createRow(0);
        String[] headers = new String[]{"Bank","Date","Balance"};
        for (int i = 0; i < headers.length; i++) {
            XSSFCell cell = (XSSFCell) row0.createCell(i);
            xssfSheet.setColumnWidth(i, 5000);
            XSSFRichTextString text = new XSSFRichTextString(headers[i]);
            cell.setCellValue(text);
        }
        int rowNum = 1;
        for (MiddleEastAll middleEastAll : MiddleEastALLList){
            XSSFRow row = xssfSheet.createRow(rowNum);
            row.createCell(0).setCellValue(middleEastAll.getBank());
            row.createCell(1).setCellValue(middleEastAll.getDocumentDate());
            row.createCell(2).setCellValue(middleEastAll.getBalance());
            rowNum++;
        }
        FileOutputStream fileOutputStream = new FileOutputStream(filePath);
        xssfWorkbook.write(fileOutputStream);

    }
}
