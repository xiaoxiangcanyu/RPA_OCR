package HttpUtil;

import DataClean.ClearAccount;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.text.DecimalFormat;
import java.util.*;

/**
 * 清账详情操作类
 */
public class MiddleEastClearData {
    public static void main(String[] args) {
//        String filePath =  args[0];
        String filePath = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\中东AR\\test\\201906041124\\AR_Clearing_List.xlsx";
        try {
            //获取清账数据表数据
            List<ClearAccount> accountList = getAccountExcelData(filePath);
            //计算手续费总额与订单总额的百分比

            double value = HandleDate(accountList).get(0).toString().equals("success")?Double.parseDouble(HandleDate(accountList).get(1).toString()):-99999999999999999.999;
            if (value!=-99999999999999999.999){
                //若手续费总额与订单总额的百分比 > 0.3% 为该数据单清账不通过
                System.out.println("value的值："+value);
                if (value - 0.3 >= 0){
                    System.out.println("大于0.3%");
                }else {
                    System.out.println("小于0.3%");
                    String filePath1 = filePath.substring(0,filePath.lastIndexOf("\\") + 1) + "success.txt";
                    File file = new File(filePath1);
                    file.createNewFile();
                }
            }else {
                String filePath1 = filePath.substring(0,filePath.lastIndexOf("\\") + 1) + "failure.txt";
                File file = new File(filePath1);
                file.createNewFile();
                if (file.exists()){
                    FileWriter fileWriter = new FileWriter(file);
                    fileWriter.write("清账模板格式不正确");
                    fileWriter.flush();
                    fileWriter.close();
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 数据获取
     * @param filePath
     * @return
     * @throws Exception
     */
    public static List getAccountExcelData(String filePath) throws Exception {
        String id = "";
        String account = "";
        String reference = "";
        String documentNo = "";
        String type = "";
        String pstngDate = "";
        String docDate = "";
        String netDueDt = "";
        String payT = "";
        String curr = "";
        String glAmount = "";
        String lCurr = "";
        String lCamnt = "";
        String clrngDoc = "";
        String gl = "";
        String text = "";
        String billDoc = "";
        String variance = "";
        String varianceMark = "";
        String costCenter = "";
        String chargeAmount = "";
        String paymentDiffAmount = "";
        File excelFile_cost = new File(filePath);
        FileInputStream EIL_file_IO = new FileInputStream(excelFile_cost);
        Workbook wb =  WorkbookFactory.create(EIL_file_IO);
        Sheet sheet = wb.getSheetAt(0);
        List<ClearAccount> accountList = new ArrayList<>();
        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row rows = sheet.getRow(r);//获取第r行
//            System.out.println("运行到第"+r+"行");
            ClearAccount clearAccount = new ClearAccount();
            try{
                rows.getCell(0);
                if (rows.getCell(0) != null){
                    rows.getCell(0).setCellType(Cell.CELL_TYPE_STRING);
                    id = rows.getCell(0).getStringCellValue();
                }else {
                    id = "";
                }
                if (rows.getCell(1) != null){
                    rows.getCell(1).setCellType(Cell.CELL_TYPE_STRING);
                    account = rows.getCell(1).getStringCellValue();
                }else {
                    account = "";
                }
                if (rows.getCell(2) != null){
                    rows.getCell(2).setCellType(Cell.CELL_TYPE_STRING);
                    reference = rows.getCell(2).getStringCellValue();
                }else {
                    reference = "";
                }
                if (rows.getCell(3) != null){
                    rows.getCell(3).setCellType(Cell.CELL_TYPE_STRING);
                    documentNo = rows.getCell(3).getStringCellValue();
                }else {
                    documentNo = "";
                }
                if (rows.getCell(4) != null){
                    rows.getCell(4).setCellType(Cell.CELL_TYPE_STRING);
                    type = rows.getCell(4).getStringCellValue();
                }else {
                    type = "";
                }
                if (rows.getCell(5) != null){
                    rows.getCell(5).setCellType(Cell.CELL_TYPE_STRING);
                    pstngDate = rows.getCell(5).getStringCellValue();
                }else {
                    pstngDate = "";
                }
                if (rows.getCell(6) != null){
                    rows.getCell(6).setCellType(Cell.CELL_TYPE_STRING);
                    docDate = rows.getCell(6).getStringCellValue();
                }else {
                    docDate = "";
                }
                if (rows.getCell(7) != null){
                    rows.getCell(7).setCellType(Cell.CELL_TYPE_STRING);
                    netDueDt = rows.getCell(7).getStringCellValue();
                }else {
                    netDueDt = "";
                }
                if (rows.getCell(8) != null){
                    rows.getCell(8).setCellType(Cell.CELL_TYPE_STRING);
                    payT = rows.getCell(8).getStringCellValue();
                }else {
                    payT = "";
                }
                if (rows.getCell(9) != null){
                    rows.getCell(9).setCellType(Cell.CELL_TYPE_STRING);
                    curr = rows.getCell(9).getStringCellValue();
                }else {
                    curr = "";
                }
                if (rows.getCell(10) != null){
                    rows.getCell(10).setCellType(Cell.CELL_TYPE_STRING);
                    glAmount = rows.getCell(10).getStringCellValue();
                }else {
                    glAmount = "";
                }
                if (rows.getCell(11) != null){
                    rows.getCell(11).setCellType(Cell.CELL_TYPE_STRING);
                    lCurr = rows.getCell(11).getStringCellValue();
                }else {
                    lCurr = "";
                }
                if (rows.getCell(12) != null){
                    rows.getCell(12).setCellType(Cell.CELL_TYPE_STRING);
                    lCamnt = rows.getCell(12).getStringCellValue();
                }else {
                    lCamnt = "";
                }
                System.out.println("在读取的时候lcamount:"+lCamnt);
                if (rows.getCell(13) != null){
                    rows.getCell(13).setCellType(Cell.CELL_TYPE_STRING);
                    clrngDoc = rows.getCell(13).getStringCellValue();
                }else {
                    clrngDoc = "";
                }
                if (rows.getCell(14) != null){
                    rows.getCell(14).setCellType(Cell.CELL_TYPE_STRING);
                    gl = rows.getCell(14).getStringCellValue();
                }else {
                    gl = "";
                }
                if (rows.getCell(15) != null){
                    rows.getCell(15).setCellType(Cell.CELL_TYPE_STRING);
                    text = rows.getCell(15).getStringCellValue();
                }else {
                    text = "";
                }
                if (rows.getCell(16) != null){
                    rows.getCell(16).setCellType(Cell.CELL_TYPE_STRING);
                    billDoc = rows.getCell(16).getStringCellValue();
                }else {
                    billDoc = "";
                }
//            System.out.println("在读取的时候va:"+rows.getCell(17));
                if (rows.getCell(17) != null){
                    rows.getCell(17).setCellType(Cell.CELL_TYPE_STRING);
                    variance = rows.getCell(17).getStringCellValue();
                }else {
                    variance = "";
                }
                if (rows.getCell(18) != null){
                    rows.getCell(18).setCellType(Cell.CELL_TYPE_STRING);
                    varianceMark = rows.getCell(18).getStringCellValue();
                }else {
                    varianceMark = "";
                }
                if (rows.getCell(19) != null){
                    rows.getCell(19).setCellType(Cell.CELL_TYPE_STRING);
                    costCenter = rows.getCell(19).getStringCellValue();
                }else {
                    costCenter = "";
                }
                if (rows.getCell(20) != null){
                    rows.getCell(20).setCellType(Cell.CELL_TYPE_STRING);
                    chargeAmount = rows.getCell(20).getStringCellValue();
                }else {
                    chargeAmount = "";
                }
                if (rows.getCell(21) != null){
                    rows.getCell(21).setCellType(Cell.CELL_TYPE_STRING);
                    paymentDiffAmount = rows.getCell(21).getStringCellValue();
                }else {
                    paymentDiffAmount = "";
                }
                clearAccount.setId(id);
                clearAccount.setAccount(account);
                clearAccount.setReference(reference);
                clearAccount.setDocumentNo(documentNo);
                clearAccount.setType(type);
                clearAccount.setPstngDate(pstngDate);
                clearAccount.setDocDate(docDate);
                clearAccount.setNetDueDt(netDueDt);
                clearAccount.setPayT(payT);
                clearAccount.setCurr(curr);
                clearAccount.setGlAmount(glAmount);
                clearAccount.setlCurr(lCurr);
                clearAccount.setlCamnt(lCamnt);
                clearAccount.setClrngDoc(clrngDoc);
                clearAccount.setGl(gl);
                clearAccount.setText(text);
                clearAccount.setBillDoc(billDoc);
                clearAccount.setVariance(variance);
//            System.out.println("在添加的时候va:"+clearAccount.getVariance());
                clearAccount.setVarianceMark(varianceMark);
                clearAccount.setCostCenter(costCenter);
                clearAccount.setChargeAmount(chargeAmount);
                clearAccount.setPaymentDiffAmount(paymentDiffAmount);
                if (!"".equals(clearAccount.getId())){
                    accountList.add(clearAccount);
                }
            }catch (Exception e){
                System.out.println("第"+r+"行为空,不计入数据");
            }
        }
        return accountList;
    }

    /**
     * 数据处理
     * @param accountList
     * @throws Exception
     */
    public static List HandleDate(List<ClearAccount> accountList){
        DecimalFormat df = new DecimalFormat("#.00");
        List<String> ResultList = new ArrayList<>();

        Set set = new HashSet();
        Map<String ,List<ClearAccount>> map = new HashMap<>();
        double total = 0;//订单总额
//        对不能按照格式读取的数据进行异常返回
        int count = 0;
        String reg = "(A0[0-9]{1})";
        for (ClearAccount clearAccount : accountList){
            if (clearAccount.getVariance().matches(reg)) {
                count++;
            }
        }
        if (count==0){
            ResultList.add("failure");
            return ResultList;
        }
        for (ClearAccount clearAccount : accountList){

            //将所有数据中金额为正数的相加，得到订单总额
            try {
                if (!clearAccount.getlCamnt().contains("-")){
                    System.out.println(clearAccount.getlCamnt());
                    total = total + Double.parseDouble(clearAccount.getlCamnt());
                }
            }catch (NumberFormatException e){
                continue;
            }

             //将数据按id分组
            if (set.add(clearAccount.getId())){
                List<ClearAccount> accountLists = new ArrayList<>();
                for (ClearAccount clearAccount1 : accountList){
                    if (clearAccount.getId().equals(clearAccount1.getId())){
                        accountLists.add(clearAccount1);
                    }
                }
                map.put(clearAccount.getId(),accountLists);
            }
        }
        double sumNum = 0;//手续费总额
        for (String key : map.keySet()) {
            List<ClearAccount> list = map.get(key);
//            int i =0;
//            for (ClearAccount account : list){
//                System.out.println("list中"+"第"+i+"个"+account.getVariance());
//                i++;
//            }
//            System.out.println("getVariance"+list.get(0).getVariance());
            //若数据Variance字段为A01 正数负数都相加
//            System.out.println(list);
//            System.out.println("代码编号:"+list.get(0).getVariance());
            if ("A01".equals(list.get(0).getVariance())){
                for (ClearAccount account : list){
//                    System.out.println("LCAmount:"+account.getlCamnt());
                    if (!account.getlCamnt().contains("-")){
                        System.out.println("正数account："+account.getlCamnt());
                        sumNum = sumNum + Double.parseDouble(account.getlCamnt());
                    }else {
                        System.out.println("负数account："+account.getlCamnt());
                        sumNum = sumNum - Double.parseDouble(account.getlCamnt().replace("-",""));
                    }

                }
            }
//            System.out.println("加完之后的LCamount:"+sumNum);
            //若数据Variance字段为A04/A05 取ChargeAmount字段数据加到手续费总额
            if ("A04".equals(list.get(0).getVariance()) || "A05".equals(list.get(0).getVariance())){
                sumNum = sumNum + Double.parseDouble(list.get(0).getChargeAmount());
            }
        }
        sumNum = Double.parseDouble(df.format(sumNum));
        total = Double.parseDouble(df.format(total));
        //计算手续费总额与订单总额的百分比
        double value = sumNum /total * 100 ;
        System.out.println("输出手续费总额：" + sumNum);
        System.out.println("输出订单总额：" + total);
        System.out.println("输出结果：" +value);
        //返回最后结果
        ResultList.add("success");
        ResultList.add(df.format(value));
//        return Double.parseDouble(df.format(value));
        return ResultList;
    }

}