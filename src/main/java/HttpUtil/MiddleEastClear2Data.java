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
public class MiddleEastClear2Data {
    public static void main(String[] args) {
        String filePath =  args[0];
//        String filePath = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\中东AR\\test\\清账20190405\\AR_Clearing_List1.xlsx";
        double value = 0;
        try {
            //获取清账数据表数据
            List<ClearAccount> accountList = getAccountExcelData(filePath);
            if (accountList.size()==0){
                 value = -99999999999999999.999;
            }else {
                value = HandleDate(accountList).get(0).toString().equals("success")?Double.parseDouble(HandleDate(accountList).get(1).toString()):-99999999999999999.999;
            }
            //计算手续费总额与订单总额的百分比

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
//        获取列数
        int columnNum = sheet.getRow(0).getPhysicalNumberOfCells();
        Row row = sheet.getRow(0);
//        检验列名是否完整
        int columnCount  = 0;
        for (int i = 0;i<columnNum;i++){
            if ("ID".equals(row.getCell(i).getStringCellValue())){
                columnCount++;
            }
            if ("Account".equals(row.getCell(i).getStringCellValue())){
                columnCount++;
            }
            if ("Reference".equals(row.getCell(i).getStringCellValue())){
                columnCount++;
            }
            if ("DocumentNo".equals(row.getCell(i).getStringCellValue())){
                columnCount++;
            }
            if ("Type".equals(row.getCell(i).getStringCellValue())){
                columnCount++;
            }
            if ("Pstng Date".equals(row.getCell(i).getStringCellValue())){
                columnCount++;
            }
            if ("Doc. Date".equals(row.getCell(i).getStringCellValue())){
                columnCount++;
            }
            if ("Net due dt".equals(row.getCell(i).getStringCellValue())){
                columnCount++;
            }
            if ("PayT".equals(row.getCell(i).getStringCellValue())){
                columnCount++;
            }
            if ("Curr.".equals(row.getCell(i).getStringCellValue())){
                columnCount++;
            }
            if (" G/L amount".equals(row.getCell(i).getStringCellValue())){
                columnCount++;
            }
            if ("LCurr".equals(row.getCell(i).getStringCellValue())){
                columnCount++;
            }
            if ("LC amnt".equals(row.getCell(i).getStringCellValue().trim())){
                columnCount++;
            }
            if ("Clrng doc.".equals(row.getCell(i).getStringCellValue())){
                columnCount++;
            }
            if ("G/L".equals(row.getCell(i).getStringCellValue())){
                columnCount++;
            }
            if ("Text".equals(row.getCell(i).getStringCellValue())){
                columnCount++;
            }
            if ("Bill.Doc.".equals(row.getCell(i).getStringCellValue())){
                columnCount++;
            }
            if ("Variance".equals(row.getCell(i).getStringCellValue())){
                columnCount++;
            }
            if ("Variance Mark".equals(row.getCell(i).getStringCellValue())){
                columnCount++;
            }
            if ("Cost Center".equals(row.getCell(i).getStringCellValue())){
                columnCount++;
            }
            if ("Charge Amount".equals(row.getCell(i).getStringCellValue())){
                columnCount++;
            }
            if ("Payment Diff Amount".equals(row.getCell(i).getStringCellValue())){
                columnCount++;
            }
        }
        System.out.println("columnCount:"+columnCount);
        System.out.println("columnNum:"+columnNum);
        if (columnCount!=22){
            List<ClearAccount> accountList = new ArrayList<>();
            return accountList;
        }
        List<ClearAccount> accountList = new ArrayList<>();
        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row rows = sheet.getRow(r);//获取第r行
//            System.out.println("运行到第"+r+"行");
            ClearAccount clearAccount = new ClearAccount();
            try{
            for (int i = 0;i<columnNum;i++){
                row.getCell(i).setCellType(Cell.CELL_TYPE_STRING);
                if (rows.getCell(i)!=null){
                    rows.getCell(i).setCellType(Cell.CELL_TYPE_STRING);
                    if ("ID".equals(row.getCell(i).getStringCellValue())){
                        id = rows.getCell(i).getStringCellValue();
                        clearAccount.setId(id);
//                        System.out.println("ID:"+id);
                    }else {
                        id = "";
                    }
                    if ("Account".equals(row.getCell(i).getStringCellValue())){
                        account = rows.getCell(i).getStringCellValue();
                        clearAccount.setAccount(account);
//                        System.out.println("account:"+account);
                    }else {
                        account = "";
                    }
                    if ("Reference".equals(row.getCell(i).getStringCellValue())){
                        reference = rows.getCell(i).getStringCellValue();
                        clearAccount.setReference(reference);
//                        System.out.println("reference:"+reference);
                    }else {
                        reference = "";
                    }
                    if ("DocumentNo".equals(row.getCell(i).getStringCellValue())){
                        documentNo = rows.getCell(i).getStringCellValue();
                        clearAccount.setDocumentNo(documentNo);
//                        System.out.println("documentNo:"+documentNo);
                    }else {
                        documentNo = "";
                    }
                    if ("Type".equals(row.getCell(i).getStringCellValue())){
                        type = rows.getCell(i).getStringCellValue();
                        clearAccount.setType(type);
//                        System.out.println("type:"+type);
                    }else {
                        type = "";
                    }
                    if ("Pstng Date".equals(row.getCell(i).getStringCellValue())){
                        pstngDate = rows.getCell(i).getStringCellValue();
                        clearAccount.setPstngDate(pstngDate);
//                        System.out.println("pstngDate:"+pstngDate);
                    }else {
                        pstngDate = "";
                    }
                    if ("Doc. Date".equals(row.getCell(i).getStringCellValue())){
                        docDate = rows.getCell(i).getStringCellValue();
                        clearAccount.setDocDate(docDate);
//                        System.out.println("docDate:"+docDate);
                    }else {
                        docDate = "";
                    }
                    if ("Net due dt".equals(row.getCell(i).getStringCellValue())){
                        netDueDt = rows.getCell(i).getStringCellValue();
                        clearAccount.setNetDueDt(netDueDt);
//                        System.out.println("netDueDt:"+netDueDt);
                    }else {
                        netDueDt = "";
                    }
                    if ("PayT".equals(row.getCell(i).getStringCellValue())){
                        payT = rows.getCell(i).getStringCellValue();
                        clearAccount.setPayT(payT);
//                        System.out.println("payT:"+payT);
                    }else {
                        payT = "";
                    }
                    if ("Curr.".equals(row.getCell(i).getStringCellValue())){
                        curr = rows.getCell(i).getStringCellValue();
                        clearAccount.setCurr(curr);
//                        System.out.println("curr:"+curr);
                    }else {
                        curr = "";
                    }
                    if (" G/L amount".equals(row.getCell(i).getStringCellValue())){
                        glAmount = rows.getCell(i).getStringCellValue();
                        clearAccount.setGlAmount(glAmount);
//                        System.out.println("glAmount:"+glAmount);
                    }else {
                        glAmount = "";
                    }
                    if ("LCurr".equals(row.getCell(i).getStringCellValue())){
                        lCurr = rows.getCell(i).getStringCellValue();
                        clearAccount.setlCurr(lCurr);
//                        System.out.println("lCurr:"+lCurr);
                    }else {
                        lCurr = "";
                    }
                    if ("LC amnt".equals(row.getCell(i).getStringCellValue().trim())){
                        lCamnt = rows.getCell(i).getStringCellValue();
                        clearAccount.setlCamnt(lCamnt);
//                        System.out.println("lCamnt:"+lCamnt);
                    }else {
                        lCamnt = "";
                    }
                    if ("Clrng doc.".equals(row.getCell(i).getStringCellValue())){
                        clrngDoc = rows.getCell(i).getStringCellValue();
                        clearAccount.setClrngDoc(clrngDoc);
//                        System.out.println("clrngDoc:"+clrngDoc);
                    }else {
                        clrngDoc = "";
                    }
                    if ("G/L".equals(row.getCell(i).getStringCellValue())){
                        gl = rows.getCell(i).getStringCellValue();
                        clearAccount.setGl(gl);
//                        System.out.println("gl:"+gl);
                    }else {
                        gl = "";
                    }
                    if ("Text".equals(row.getCell(i).getStringCellValue())){
                        text = rows.getCell(i).getStringCellValue();
                        clearAccount.setText(text);
//                        System.out.println("text:"+text);
                    }else {
                        text = "";
                    }
                    if ("Bill.Doc.".equals(row.getCell(i).getStringCellValue())){
                        billDoc = rows.getCell(i).getStringCellValue();
                        clearAccount.setBillDoc(billDoc);
//                        System.out.println("billDoc:"+billDoc);
                    }else {
                        billDoc = "";
                    }
                    if ("Variance".equals(row.getCell(i).getStringCellValue())){
                        variance = rows.getCell(i).getStringCellValue();
                        clearAccount.setVariance(variance);
//                        System.out.println("variance:"+variance);
                    }else {
                        variance = "";
                    }
                    if ("Variance Mark".equals(row.getCell(i).getStringCellValue())){
                        varianceMark = rows.getCell(i).getStringCellValue();
                        clearAccount.setVarianceMark(varianceMark);
//                        System.out.println("varianceMark:"+varianceMark);
                    }else {
                        varianceMark = "";
                    }
                    if ("Cost Center".equals(row.getCell(i).getStringCellValue())){
                        costCenter = rows.getCell(i).getStringCellValue();
                        clearAccount.setCostCenter(costCenter);
//                        System.out.println("costCenter:"+costCenter);
                    }else {
                        costCenter = "";
                    }
                    if ("Charge Amount".equals(row.getCell(i).getStringCellValue())){
                        chargeAmount = rows.getCell(i).getStringCellValue();
                        clearAccount.setChargeAmount(chargeAmount);
//                        System.out.println("chargeAmount:"+chargeAmount);
                    }
                    if ("Payment Diff Amount".equals(row.getCell(i).getStringCellValue())){
                        paymentDiffAmount = rows.getCell(i).getStringCellValue();
                        clearAccount.setPaymentDiffAmount(paymentDiffAmount);
//                        System.out.println("paymentDiffAmount:"+paymentDiffAmount);
                    }else {
                        paymentDiffAmount = "";
                    }

//            System.out.println("在添加的时候va:"+clearAccount.getVariance());

                }

            }

            }catch (Exception e){
                System.out.println("第"+r+"行为空,不计入数据");
            }
            if (!"".equals(clearAccount.getId()) && clearAccount.getId()!=null){
//                System.out.println("加入集合的id:"+clearAccount.getId()+"加入集合的Reference："+clearAccount.getReference());
                accountList.add(clearAccount);
            }
        }

        System.out.println("共添加了"+accountList.size()+"行数据");
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
            System.out.println("clearAccount.getVariance():"+clearAccount.getVariance());
            if (clearAccount.getVariance()!=null){
                if (clearAccount.getVariance().matches(reg)) {
                    count++;
                }
            }
        }
        if (count==0){
            ResultList.add("failure");
            return ResultList;
        }
        for (ClearAccount clearAccount : accountList){
            System.out.println(clearAccount.getlCamnt());
            //将所有数据中金额为正数的相加，得到订单总额
            if (!"".equals(clearAccount.getlCamnt()) && !clearAccount.getlCamnt().contains("-")){
                total = total + Double.parseDouble(clearAccount.getlCamnt());
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
                    if (!"".equals(account.getlCamnt())){
                        if (!account.getlCamnt().contains("-")){
                            System.out.println("正数account："+account.getlCamnt());
                            sumNum = sumNum + Double.parseDouble(account.getlCamnt());
                        }else {
                            System.out.println("负数account："+account.getlCamnt());
                            sumNum = sumNum - Double.parseDouble(account.getlCamnt().replace("-",""));
                        }
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