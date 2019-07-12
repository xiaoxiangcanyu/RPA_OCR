package HttpUtil;

import DataClean.BaseUtil;
import DataClean.MiddleEastAll;
import DataClean.MiddleEastSAP;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Pattern;

public class MiddleEastForOther extends BaseUtil {
    public static void main(String[] args) throws Exception{
        String filePathLedger = args[0];//全量台账数据表
        String filePathSap = args[1];//FB01表路径
        String filePath = args[2];//生成F-32表路径
//        String filePathLedger = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\中东AR\\test\\20190305\\AR_Matching_List.xlsx";//全量台账数据表
//        String filePathSap = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\中东AR\\test\\20190305\\FB01_2.xlsx";//FB01表路径
//        String filePath = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\中东AR\\test\\20190305\\F-32.xls";//生成F-32表路径
        try {
            // 获取台账数据
            List<MiddleEastAll> MiddleEastLedgerList = getLedgerExcelData(filePathLedger);
            //获取FB01(sap)表数据
            List<MiddleEastSAP> middleEastSAPList = getMiddleEastSAP(filePathSap);

            List<MiddleEastAll> MiddleEast = new ArrayList<>();//定义集合存储可生成F-32 表的数据
            for (MiddleEastAll middleEastAll : MiddleEastLedgerList){
                Pattern p = Pattern.compile("[^0-9]");//判断incomeNo中是否包含10为数字编码
                String sapIncomeNo = p.matcher(middleEastAll.getSapIncomeNo()).replaceAll("").trim();
                Set set = new HashSet();//定义set来去重数据
                String str = "";
                //用来判断该条数据是否有对应的FB01数据
                boolean flag = false;
                if (sapIncomeNo.length() < 10 ){
                    //将FB01 表中ID相同数据的 “FB01 Number”字段去重拼接
                    for (MiddleEastSAP middleEastSAP : middleEastSAPList){
//                        System.out.println("全量台账的id:"+middleEastAll.getId()+",FB01的id:"+middleEastSAP.getId());
                        if (middleEastAll.getId().equals(middleEastSAP.getId())){
//                            System.out.println("id相同:"+middleEastSAP.getId());
                            flag = true;
                            if (set.add(middleEastSAP.getFbNumber())){
                                str = str + middleEastSAP.getFbNumber() + " ";
//                                System.out.println("str:"+str);
                            }
                        }
                    }
                    //判断当前台账数据是否是一次性客户类型数据，回写对应字段
                    if (flag){
                        if ("other".equals(middleEastAll.getSapIncomeNo())){
                            if (!"".equals(str.trim())){
                                middleEastAll.setSapNoForOther(str.trim());
                            }else {
                                middleEastAll.setSapNoForOther("Error In SAP");
                            }
                        }else {
                            if (!"".equals(middleEastAll.getCustomerCode())){
                                if (!"".equals(str.trim())){
                                    middleEastAll.setSapIncomeNo(str.trim());
                                    System.out.println("id:"+middleEastAll.getId()+"走的这里1");
                                }else {
                                    System.out.println("id:"+middleEastAll.getId()+"走的这里2");
                                    middleEastAll.setSapIncomeNo("Error In SAP");
                                }
                            }else {
                                middleEastAll.setSapIncomeNo("");
                            }
                        }
                    }else {
                        if ("".equals(middleEastAll.getCustomerCode())){
                            middleEastAll.setSapIncomeNo("");
                        }
                    }
                    Pattern pp = Pattern.compile("[^0-9]");//判断incomeNo中是否包含10为数字编码
                    System.out.println("在正则校验之前的sapincomeNum："+middleEastAll.getSapIncomeNo()+",id:"+middleEastAll.getId());
                    String incomeNo = pp.matcher(middleEastAll.getSapIncomeNo()).replaceAll("").trim();
                    //判断数据是否为一次性客户类型数据认领，是则生成F-32表
                    System.out.println("incomeNo:"+incomeNo.length() );
//                    System.out.println("SapNoForOther："+middleEastAll.getSapNoForOther());
//                    System.out.println("middleEastAll.getSapClearingNo():"+middleEastAll.getSapClearingNo());
//                    System.out.println("middleEastAll.getIncome():"+middleEastAll.getIncome());
                    if (incomeNo.length() >= 10 && !"".equals(middleEastAll.getSapNoForOther()) && "".equals(middleEastAll.getSapClearingNo())
                            && (!"".equals(middleEastAll.getIncome()) && middleEastAll.getIncome() != null)){
                        System.out.println("符合生成F-32条件");
                        MiddleEast.add(middleEastAll);
                    }
                }
            }
            //回写台账
            excelOutput_Log(MiddleEastLedgerList, filePathLedger);
            if (MiddleEast.size() > 0){
                //生成F-32表
                excelOutputFB(MiddleEast,filePath);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 获取FB01(SAP)表数据
     * @param filePath
     * @return
     * @throws Exception
     */
    public static List getMiddleEastSAP(String filePath) throws Exception {
        File excelFile = new File(filePath);
        FileInputStream EIL_file_IO = new FileInputStream(excelFile);
        Workbook wb = WorkbookFactory.create(EIL_file_IO);
        Sheet sheet = wb.getSheetAt(0);
        //获取列数
        int coloumNum = sheet.getRow(0).getPhysicalNumberOfCells();
        //获取表头行数据
        Row row = sheet.getRow(0);
        List<MiddleEastSAP> MiddleEastSapList = new ArrayList<>();
        for (int r = 1; r <= sheet.getLastRowNum(); r++){
            String id = "";
            String fbNumber = "";
            MiddleEastSAP middleEastSAP = new MiddleEastSAP();
            Row rows = sheet.getRow(r);//获取第r行
            for (int i = 0; i < coloumNum ; i ++){
                if (rows.getCell(i) != null){
                    row.getCell(i).setCellType(Cell.CELL_TYPE_STRING);
                    rows.getCell(i).setCellType(Cell.CELL_TYPE_STRING);
                    if ("ID".equals(row.getCell(i).getStringCellValue())){
                        id = rows.getCell(i).getStringCellValue();
                        System.out.println("ID:"+id);
                    }
                    if ("FB01 Number".equals(row.getCell(i).getStringCellValue())){
                        fbNumber = rows.getCell(i).getStringCellValue();
                    }
                }
            }
//            System.out.println("FB01Number:"+fbNumber);
            middleEastSAP.setId(id);
            middleEastSAP.setFbNumber(fbNumber);
            if (!"".equals(id)){
                MiddleEastSapList.add(middleEastSAP);
            }
        }
        return MiddleEastSapList;
    }

    /**
     * 生成F-32 数据表
     * @param MiddleEast
     * @param filePath
     * @throws Exception
     */
    public static void excelOutputFB( List<MiddleEastAll> MiddleEast, String filePath) throws Exception{
        //创建表格
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        //定义第一个sheet页
        XSSFSheet xssfSheet = xssfWorkbook.createSheet("Sheet1");
        //第一个sheet页数据（生成台账表）
        Row row0 = xssfSheet.createRow(0);
        String[] headers = new String[]{"ID","No","Account", "Clearing Date", "Period", "Company Code", "Currency", "SGL Ind", "Document No.", "Reference","EmailAddress"};
        for (int i = 0; i < headers.length; i++) {
            XSSFCell cell = (XSSFCell) row0.createCell(i);
            XSSFRichTextString text = new XSSFRichTextString(headers[i]);
            xssfSheet.setColumnWidth(i, 5000);
            cell.setCellValue(text);
        }
        int rowNum = 1;
        for (MiddleEastAll middleEastAll :MiddleEast){
            if ("TT".equals(middleEastAll.getTtLCMark()) || "LC".equals(middleEastAll.getTtLCMark())){
                XSSFRow row = xssfSheet.createRow(rowNum);
                xssfSheet.setColumnWidth(rowNum, 5000);
                String uuid = UUID.randomUUID().toString().replaceAll("-","");
                row.createCell(0).setCellValue(middleEastAll.getId());
                row.createCell(1).setCellValue(uuid);
                row.createCell(2).setCellValue("6000007049");
                DateFormat dFormat = new SimpleDateFormat("MM/dd/yyyy");
                Calendar calendar = Calendar.getInstance();
                calendar.setTime(new Date());
                calendar.set(Calendar.DAY_OF_MONTH,1);//当前日期既为本月第一天
                row.createCell(3).setCellValue(dFormat.format(calendar.getTime()));
                String  month = String.valueOf(calendar.get(Calendar.MONTH) + 1);
                row.createCell(4).setCellValue(month);
                row.createCell(5).setCellValue("6670");
                row.createCell(6).setCellValue(middleEastAll.getCurrency());
                row.createCell(7).setCellValue("A");
                row.createCell(8).setCellValue(middleEastAll.getSapIncomeNo());

                XSSFRow row2 = xssfSheet.createRow(rowNum + 1);
                xssfSheet.setColumnWidth(rowNum + 1, 5000);
                String uuid2 = UUID.randomUUID().toString().replaceAll("-","");
                row2.createCell(0).setCellValue(middleEastAll.getId());
                row2.createCell(1).setCellValue(uuid2);
                row2.createCell(2).setCellValue("6000007049");
                row2.createCell(3).setCellValue(dFormat.format(calendar.getTime()));
                row2.createCell(4).setCellValue(month);
                row2.createCell(5).setCellValue("6670");
                row2.createCell(6).setCellValue(middleEastAll.getCurrency());
                row2.createCell(7).setCellValue("A");
                row2.createCell(8).setCellValue(middleEastAll.getSapNoForOther());
                row2.createCell(9).setCellValue(middleEastAll.getEmail());

                rowNum+=2;
            }
        }
        FileOutputStream fileOutputStream = new FileOutputStream(filePath);
        xssfWorkbook.write(fileOutputStream);
    }
}
