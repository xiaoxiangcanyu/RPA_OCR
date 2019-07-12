package HttpUtil;

import DataClean.BaseUtil;
import DataClean.CustomerCodeDO;
import DataClean.MiddleEastAll;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.regex.Pattern;

public class MiddleEast_Combine extends BaseUtil {
    public static void main(String[] args) {
        String filePath_New = args[0];
        String filePath_Old = args[1];
        String CustomerCodeFile = args[2];
        String EmailAddress = args[3];
//        String filePath_Old = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\中东AR\\test\\201907032309\\AR_Matching_List2019-06-04 -基础.xlsx";//全量台账数据表
//        String filePath_New = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\中东AR\\test\\201907032309\\AR_Matching_List2019-06-04-认领.xlsx";//生成未认领数据台账表路径
//        String CustomerCodeFile = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\中东AR\\test\\2019060111805\\Combine优化\\customer code list.xlsx";
//        String EmailAddress = "123@haier.com";
        try {
            //获取全量台账表数据
            List<MiddleEastAll> OldLedgerList = getLedgerExcelData(filePath_Old);
            //原台账去重
            OldLedgerList = RemoveDulplicate(OldLedgerList);
            //获取认领台账表数据
            List<MiddleEastAll> NewLedgerList = getLedgerExcelData(filePath_New);
            NewLedgerList = HandleData(NewLedgerList);
            //数据整合
            List<MiddleEastAll> MiddleEastLedgerList = getMiddleEast(OldLedgerList, NewLedgerList, EmailAddress);
            //数据整合之后去重
            MiddleEastLedgerList = RemoveDulplicate(MiddleEastLedgerList);
            //SCB-TT校验
            MiddleEastLedgerList = SCBCheck(MiddleEastLedgerList, CustomerCodeFile);
            //数据校验
            MiddleEastLedgerList = DataCheck(MiddleEastLedgerList);
//            //数据整合之后去重
            MiddleEastLedgerList = RemoveDulplicate(MiddleEastLedgerList);
            //中东临时客户认领为A01的把SAP for other号复制到 sap income no里面
            MiddleEastLedgerList = CopySAPFOrOther(MiddleEastLedgerList);
//            生成新的全量台账表
            excelOutput_Log(MiddleEastLedgerList, filePath_Old);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 中东临时客户认领为A01的把SAP for other号复制到 sap income no里面
     *
     * @param middleEastLedgerList
     */
    private static List<MiddleEastAll> CopySAPFOrOther(List<MiddleEastAll> middleEastLedgerList) {
        for (MiddleEastAll middleEastAll : middleEastLedgerList) {
            if (middleEastAll.getSapNoForOther() != null && middleEastAll.getSapNoForOther().length() > 1) {
                if ("A01".equals(middleEastAll.getCustomerCode())) {
                    middleEastAll.setSapIncomeNo(middleEastAll.getSapNoForOther());
                }
            }
        }
        return middleEastLedgerList;
    }

    /**
     * SCB校验
     *
     * @param middleEastLedgerList
     * @param customerCodeFile
     * @return
     */
    private static List<MiddleEastAll> SCBCheck(List<MiddleEastAll> middleEastLedgerList, String customerCodeFile) throws IOException, InvalidFormatException {
        List<CustomerCodeDO> customerCodeDOList = GetCustomerCodeList(customerCodeFile);
        for (MiddleEastAll middleEastAll : middleEastLedgerList) {
            if (middleEastAll.getBank().startsWith("SCB") && "TT".equals(middleEastAll.getTtLCMark())) {
                //匹配成功
                for (CustomerCodeDO customerCodeDO : customerCodeDOList) {
                    if (middleEastAll.getCustomerCode().equals(customerCodeDO.getCustomerCode())) {
                        if (middleEastAll.getSummary().contains(customerCodeDO.getName())) {
                            middleEastAll.setComment("CustomerName匹配成功!");
                        }
                    }
                }
                //匹配不成功
                if (middleEastAll.getComment() == null) {
                    middleEastAll.setComment("CustomerName未匹配成功!");
                }
            }
        }
        return middleEastLedgerList;
    }

    /**
     * 从CustomerCodeList中获取数据
     *
     * @param customerCodeFile
     * @return
     */
    private static List<CustomerCodeDO> GetCustomerCodeList(String customerCodeFile) throws IOException, InvalidFormatException {
        List<CustomerCodeDO> list = new ArrayList<>();
        File file = new File(customerCodeFile);
        FileInputStream fileInputStream = new FileInputStream(file);
        Workbook workbook = WorkbookFactory.create(fileInputStream);
        Sheet sheet = workbook.getSheet("Sheet1");
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            if (sheet != null) {
                Row row = sheet.getRow(i);
                CustomerCodeDO customerCodeDO = new CustomerCodeDO();
                if (row != null) {
                    if (row.getCell(0) != null) {
                        row.getCell(0).setCellType(Cell.CELL_TYPE_STRING);
                        customerCodeDO.setCustomerCode(row.getCell(0).getStringCellValue());
                    }
                    if (row.getCell(1) != null) {
                        row.getCell(1).setCellType(Cell.CELL_TYPE_STRING);
                        customerCodeDO.setName(row.getCell(1).getStringCellValue());
                    }
                    list.add(customerCodeDO);
                }

            }

        }
        return list;

    }

    /**
     * 原台账去重
     *
     * @param oldLedgerList
     * @return
     */
    private static List<MiddleEastAll> RemoveDulplicate(List<MiddleEastAll> oldLedgerList) {
        List<MiddleEastAll> list = new ArrayList<>();
        Set<String> set = new HashSet<>();
        for (MiddleEastAll middleEastAll : oldLedgerList) {
            String str = middleEastAll.getId() + middleEastAll.getPi() + middleEastAll.getProductCode();
            if (set.add(str)) {
                list.add(middleEastAll);
            }

        }

        return list;
    }

    /**
     * 数据整合
     * 将业务认领的台账表与原始的台账表进行数据整合
     *
     * @param OldLedgerList
     * @param NewLedgerList
     * @param EmailAddress
     * @return
     * @throws Exception
     */
    public static List getMiddleEast(List<MiddleEastAll> OldLedgerList, List<MiddleEastAll> NewLedgerList, String EmailAddress) throws Exception {
        List<MiddleEastAll> MiddleEastList = new ArrayList<>();
        Set set = new HashSet();
        Set set0 = new HashSet();
        Set set1 = new HashSet();
        Set set3 = new HashSet();
        Set set4 = new HashSet();
        Set set5 = new HashSet();
        Set set6 = new HashSet();
        System.out.println("开始处理数据。。");
        //遍历原始的台账数据
        for (MiddleEastAll OldMiddleEast : OldLedgerList) {
            int flag = 0;
            int flag2 = 0;
            int flag3 = 0;
            Set set7 = new HashSet();
            //遍历业务认领 之后的台账数据
            for (MiddleEastAll NewMiddleEast : NewLedgerList) {
                //分别提取数据的SapIncomeNo字段的内容，判断数据是否需要进行数据合并
                Pattern p = Pattern.compile("[^0-9]");
                String oldIncomeNo = p.matcher(OldMiddleEast.getSapIncomeNo()).replaceAll("").trim();
                String newIncomeNo = p.matcher(NewMiddleEast.getSapIncomeNo()).replaceAll("").trim();
                //如果旧的台账的SAPIncomeNum为"检验成功",则不需要覆盖新的
                if (!"校验成功!".equals(OldMiddleEast.getSapIncomeNo())) {
                    if (oldIncomeNo.length() < 10) {//全量台账数据中SapIncomeNo不包含10位数字编号
                        //认领数据中CustomerCode不为空以及SapIncomeNo值不包含10位数字编号
                        if (!"".equals(NewMiddleEast.getCustomerCode()) && newIncomeNo.length() < 10) {
                            //判断数据类型
                            switch (NewMiddleEast.getTtLCMark()) {
                                case "LC":
                                    if (OldMiddleEast.getId().equals(NewMiddleEast.getId())) {
                                        String str = NewMiddleEast.getId() + NewMiddleEast.getPi() + NewMiddleEast.getProductCode();
                                        if (set0.add(str)) {
                                            MiddleEastAll NewMiddleEast_cover = coverOld(OldMiddleEast, NewMiddleEast);
                                            NewMiddleEast_cover.setEmail(EmailAddress);
                                            MiddleEastList.add(NewMiddleEast_cover);
                                        }
                                    }
                                    break;
                                case "TT":
                                    if (OldMiddleEast.getId().equals(NewMiddleEast.getId())) {
                                        flag = flag + 1;
                                        //判断当前数据是收款数据
                                        if (!"".equals(OldMiddleEast.getIncome()) && "".equals(OldMiddleEast.getCharge()) && !"".equals(OldMiddleEast.getCustomerCode())) {
                                            //判断当前数据的pi以及产品线是否存在
                                            if (!"".equals(OldMiddleEast.getPi()) && !"".equals(OldMiddleEast.getProductCode())) {
//                                                System.out.println("pi以及产品线存在");
                                                if ((OldMiddleEast.getPi().equals(NewMiddleEast.getPi())) && (OldMiddleEast.getProductCode().equals(NewMiddleEast.getProductCode()))) {
                                                    if (set1.add(NewMiddleEast.getId())) {
                                                        MiddleEastAll NewMiddleEast_cover = coverOld(OldMiddleEast, NewMiddleEast);
                                                        NewMiddleEast_cover.setEmail(EmailAddress);
                                                        MiddleEastList.add(NewMiddleEast_cover);
                                                    }
                                                } else {
                                                    String str = OldMiddleEast.getId() + OldMiddleEast.getPi() + OldMiddleEast.getProductCode();
                                                    if (set.add(str)) {
                                                        set1.add(OldMiddleEast.getId());
                                                        MiddleEastList.add(OldMiddleEast);
                                                    }
                                                    String str1 = NewMiddleEast.getId() + NewMiddleEast.getPi() + NewMiddleEast.getProductCode();
                                                    if (set.add(str1)) {
                                                        MiddleEastAll NewMiddleEast_cover = coverOld(OldMiddleEast, NewMiddleEast);
                                                        NewMiddleEast_cover.setEmail(EmailAddress);
                                                        MiddleEastList.add(NewMiddleEast_cover);
                                                    } else {
                                                        if (set.add(str)) {
                                                            NewMiddleEast.setEmail(EmailAddress);
                                                            MiddleEastList.add(NewMiddleEast);
                                                        }
                                                    }
//                                                    if (flag > 1){
//                                                        //相同id所对应的每条数据的pi以及产品线是唯一。固使用id，pi以及产品线组成唯一的字符串
//
//                                                    } else {

//                                                        if (set1.add(NewMiddleEast.getId())){
//                                                            MiddleEastAll NewMiddleEast_cover = coverOld(OldMiddleEast,NewMiddleEast);
//                                                            MiddleEastList.add(OldMiddleEast);
//                                                        }
//                                                    }
                                                }
                                            } else {
                                                MiddleEastAll NewMiddleEast_cover = coverOld(OldMiddleEast, NewMiddleEast);
                                                NewMiddleEast_cover.setEmail(EmailAddress);
                                                MiddleEastList.add(NewMiddleEast_cover);
                                            }
                                        } else {
                                            MiddleEastAll NewMiddleEast_cover = coverOld(OldMiddleEast, NewMiddleEast);
                                            NewMiddleEast_cover.setEmail(EmailAddress);
                                            MiddleEastList.add(NewMiddleEast_cover);
                                        }
                                    }
                                    break;
                                case "TT/LC":
                                    if (OldMiddleEast.getId().equals(NewMiddleEast.getId())) {
                                        flag2 = flag2 + 1;
                                        //判断当前数据是收款数据
                                        if (!"".equals(OldMiddleEast.getIncome()) && "".equals(OldMiddleEast.getCharge()) && !"".equals(OldMiddleEast.getCustomerCode())) {
                                            //判断当前数据的pi以及产品线是否存在
                                            if (!"".equals(OldMiddleEast.getPi()) && !"".equals(OldMiddleEast.getProductCode())) {
                                                //判断新数据和老数据的pi以及产品线是否同时相同
                                                if ((OldMiddleEast.getPi().equals(NewMiddleEast.getPi())) && (OldMiddleEast.getProductCode().equals(NewMiddleEast.getProductCode()))) {
                                                    if (set4.add(NewMiddleEast.getId())) {
                                                        MiddleEastAll NewMiddleEast_cover = coverOld(OldMiddleEast, NewMiddleEast);
                                                        NewMiddleEast_cover.setEmail(EmailAddress);
                                                        MiddleEastList.add(NewMiddleEast_cover);
                                                    }
                                                } else {
                                                    if (flag2 > 1) {
                                                        //相同id所对应的每条数据的pi以及产品线是唯一。固使用id，pi以及产品线组成唯一的字符串
                                                        String str = NewMiddleEast.getId() + NewMiddleEast.getPi() + NewMiddleEast.getProductCode();
                                                        if (set3.add(str)) {
                                                            MiddleEastAll NewMiddleEast_cover = coverOld(OldMiddleEast, NewMiddleEast);
                                                            NewMiddleEast_cover.setEmail(EmailAddress);
                                                            MiddleEastList.add(NewMiddleEast_cover);
                                                        }
                                                    } else {
                                                        String str = OldMiddleEast.getId() + OldMiddleEast.getPi() + OldMiddleEast.getProductCode();
                                                        if (set3.add(str)) {
                                                            MiddleEastList.add(OldMiddleEast);
                                                        }
                                                        if (set4.add(NewMiddleEast.getId())) {
                                                            MiddleEastAll NewMiddleEast_cover = coverOld(OldMiddleEast, NewMiddleEast);
                                                            NewMiddleEast_cover.setEmail(EmailAddress);
                                                            MiddleEastList.add(NewMiddleEast_cover);
                                                        }
                                                    }
                                                }
                                            } else {
                                                MiddleEastAll NewMiddleEast_cover = coverOld(OldMiddleEast, NewMiddleEast);
                                                NewMiddleEast_cover.setEmail(EmailAddress);
                                                MiddleEastList.add(NewMiddleEast_cover);
                                            }
                                        } else {
                                            MiddleEastAll NewMiddleEast_cover = coverOld(OldMiddleEast, NewMiddleEast);
                                            NewMiddleEast_cover.setEmail(EmailAddress);
                                            MiddleEastList.add(NewMiddleEast_cover);
                                        }
                                    }
                                    break;
                                default:
                                    //非任何类型的数据
                                    if (OldMiddleEast.getId().equals(NewMiddleEast.getId())) {
                                        flag3 = flag3 + 1;
                                        //判断当前数据是收款数据
                                        if (!"".equals(OldMiddleEast.getIncome()) && "".equals(OldMiddleEast.getCharge()) && !"".equals(OldMiddleEast.getCustomerCode())) {
                                            //判断当前数据的pi以及产品线是否存在
                                            if (!"".equals(OldMiddleEast.getPi()) && !"".equals(OldMiddleEast.getProductCode())) {
                                                //判断新数据和老数据的pi以及产品线是否同时相同
                                                if ((OldMiddleEast.getPi().equals(NewMiddleEast.getPi())) && (OldMiddleEast.getProductCode().equals(NewMiddleEast.getProductCode()))) {
                                                    if (set5.add(NewMiddleEast.getId())) {
                                                        MiddleEastAll NewMiddleEast_cover = coverOld(OldMiddleEast, NewMiddleEast);
                                                        NewMiddleEast_cover.setEmail(EmailAddress);
                                                        MiddleEastList.add(NewMiddleEast_cover);
                                                    }
                                                } else {
                                                    if (flag3 > 1) {
                                                        //相同id所对应的每条数据的pi以及产品线是唯一。固使用id，pi以及产品线组成唯一的字符串
                                                        String str = NewMiddleEast.getId() + NewMiddleEast.getPi() + NewMiddleEast.getProductCode();
                                                        if (set6.add(str)) {
                                                            MiddleEastAll NewMiddleEast_cover = coverOld(OldMiddleEast, NewMiddleEast);
                                                            NewMiddleEast_cover.setEmail(EmailAddress);
                                                            MiddleEastList.add(NewMiddleEast_cover);
                                                        }
                                                    } else {
                                                        String str = OldMiddleEast.getId() + OldMiddleEast.getPi() + OldMiddleEast.getProductCode();
                                                        if (set6.add(str)) {
                                                            MiddleEastList.add(OldMiddleEast);
                                                        }
                                                        if (set5.add(NewMiddleEast.getId())) {
                                                            MiddleEastAll NewMiddleEast_cover = coverOld(OldMiddleEast, NewMiddleEast);
                                                            NewMiddleEast_cover.setEmail(EmailAddress);
                                                            MiddleEastList.add(NewMiddleEast_cover);
                                                        }
                                                    }
                                                }
                                            } else {
                                                MiddleEastAll NewMiddleEast_cover = coverOld(OldMiddleEast, NewMiddleEast);
                                                NewMiddleEast_cover.setEmail(EmailAddress);
                                                MiddleEastList.add(NewMiddleEast_cover);
                                            }
                                        } else {
                                            MiddleEastAll NewMiddleEast_cover = coverOld(OldMiddleEast, NewMiddleEast);
                                            NewMiddleEast_cover.setEmail(EmailAddress);
                                            MiddleEastList.add(NewMiddleEast_cover);
                                        }
                                    }
                                    break;
                            }
                        } else {
                            if (OldMiddleEast.getId().equals(NewMiddleEast.getId())) {
                                MiddleEastList.add(OldMiddleEast);
                            }
                        }
                    } else {
                        if (set7.add(NewMiddleEast.getId())) {
                            if (OldMiddleEast.getId().equals(NewMiddleEast.getId())) {
                                MiddleEastList.add(OldMiddleEast);
                            }
                        }
                    }
                } else {
                    MiddleEastList.add(OldMiddleEast);
                }

            }


            int index = 0;
            for (MiddleEastAll NewMiddleEast : NewLedgerList) {
                if (OldMiddleEast.getId().equals(NewMiddleEast.getId())) {
                    index++;
                }
            }
            if (index == 0) {
                MiddleEastList.add(OldMiddleEast);
            }
        }
        System.out.println("处理数据结束。。");
        return MiddleEastList;
    }

    /**
     * 覆盖新的台账
     *
     * @param oldmiddleeast
     * @param newmiddleeast
     * @return
     */
    private static MiddleEastAll coverOld(MiddleEastAll oldmiddleeast, MiddleEastAll newmiddleeast) {
        MiddleEastAll middleEastAllTemp = new MiddleEastAll();
        middleEastAllTemp.setPi(oldmiddleeast.getPi());
        middleEastAllTemp.setRemark(oldmiddleeast.getRemark());
        middleEastAllTemp.setStaffName(oldmiddleeast.getStaffName());
        middleEastAllTemp.setDepositInterest(oldmiddleeast.getDepositInterest());
        middleEastAllTemp.setDepositPrincipal(oldmiddleeast.getDepositPrincipal());
        middleEastAllTemp.setTransferTo(oldmiddleeast.getTransferTo());
        middleEastAllTemp.setPrepayment(oldmiddleeast.getPrepayment());
        middleEastAllTemp.setRecognizedAmount(oldmiddleeast.getRecognizedAmount());
        middleEastAllTemp.setCustomerName(oldmiddleeast.getCustomerName());
        middleEastAllTemp.setCustomerCode(oldmiddleeast.getCustomerCode());
        middleEastAllTemp.setSapIncomeNo(oldmiddleeast.getSapIncomeNo());
        middleEastAllTemp.setTtLCMark(oldmiddleeast.getTtLCMark());
        middleEastAllTemp.setBank(oldmiddleeast.getBank());
        middleEastAllTemp.setDocumentDate(oldmiddleeast.getDocumentDate());
        middleEastAllTemp.setCurrency(oldmiddleeast.getCurrency());
        middleEastAllTemp.setBalance(oldmiddleeast.getBalance());
        middleEastAllTemp.setCharge(oldmiddleeast.getCharge());
        middleEastAllTemp.setIncome(oldmiddleeast.getIncome());
        middleEastAllTemp.setSummary(oldmiddleeast.getSummary());
        middleEastAllTemp.setProductCode(oldmiddleeast.getProductCode());
        middleEastAllTemp.setSapClearingNo(oldmiddleeast.getSapClearingNo());
        middleEastAllTemp.setId(oldmiddleeast.getId());
        middleEastAllTemp.setSapNoForOther(oldmiddleeast.getSapNoForOther());

        middleEastAllTemp.setCustomerCode(newmiddleeast.getCustomerCode());
        middleEastAllTemp.setCustomerName(newmiddleeast.getCustomerName());
        middleEastAllTemp.setRecognizedAmount(newmiddleeast.getRecognizedAmount());
        middleEastAllTemp.setProductCode(newmiddleeast.getProductCode());
        middleEastAllTemp.setPi(newmiddleeast.getPi());
        middleEastAllTemp.setPrepayment(newmiddleeast.getPrepayment());
        middleEastAllTemp.setTransferTo(newmiddleeast.getTransferTo());
        middleEastAllTemp.setDepositPrincipal(newmiddleeast.getDepositPrincipal());
        middleEastAllTemp.setDepositInterest(newmiddleeast.getDepositInterest());
        middleEastAllTemp.setStaffName(newmiddleeast.getStaffName());
        middleEastAllTemp.setRemark(newmiddleeast.getRemark());
        return middleEastAllTemp;
    }

    /**
     * 数据校验
     *
     * @param middleEastLedgerList
     * @return
     */
    private static List<MiddleEastAll> DataCheck(List<MiddleEastAll> middleEastLedgerList) throws Exception {

        DecimalFormat df = new DecimalFormat("#.00");
        for (MiddleEastAll middleEastAll : middleEastLedgerList) {
            Pattern pp = Pattern.compile("[^0-9]");//判断sapIncomeNo中是否包含10位数字编码
            String incomeNo = pp.matcher(middleEastAll.getSapIncomeNo()).replaceAll("").trim();
            String customerCode = middleEastAll.getCustomerCode();
            String status = "";
            if (incomeNo.length() < 10) {
                if (customerCode != null && !"".equals(customerCode)) {
                    if (!(middleEastAll.getStaffName() != null && !"".equals(middleEastAll.getStaffName()))) {
                        status = status + "未填写StaffName!";
                    }
                }
                middleEastAll.setSapIncomeNo(status);
                if (customerCode.contains("6000") && customerCode.indexOf("6000", 0) == 0 && customerCode.length() == 10) {
                    status = middleEastAll.getSapIncomeNo();
                    //            ================================================校验1==========================================================================================
                    if ("".equals(middleEastAll.getCustomerName()) || middleEastAll.getCustomerName() == null) {
                        status = status + "客户名称为空!";
                    }
                    if ("".equals(middleEastAll.getCustomerCode()) || middleEastAll.getCustomerCode() == null) {
                        status = status + "客户代码为空!";
                    }
                    if ("".equals(middleEastAll.getRecognizedAmount()) || middleEastAll.getRecognizedAmount() == null) {
                        status = status + "认领金额为空!";
                    }
                    if ("".equals(middleEastAll.getPi()) || middleEastAll.getPi() == null) {
                        status = status + "PI为空!";
                    }
                    if ("".equals(middleEastAll.getProductCode()) || middleEastAll.getProductCode() == null) {
                        status = status + "产品线为空!";
                    }
                    if ("".equals(middleEastAll.getPrepayment()) || middleEastAll.getPrepayment() == null) {
                        status = status + "预付比例为空!";
                    } else {
                        if (!middleEastAll.getPrepayment().contains("N") && !middleEastAll.getPrepayment().contains("Y")
                                && !middleEastAll.getPrepayment().contains("尾款")) {
                            status = status + "预付比例不符合规则!";
                        }
                    }
                    middleEastAll.setSapIncomeNo(status);
//            ================================================校验1==========================================================================================
//            ================================================校验2==========================================================================================
                    status = middleEastAll.getSapIncomeNo();
                    Double sum = 0.0;
                    Double income = 0.0;
                    if (middleEastAll.getIncome() != null && middleEastAll.getIncome().length() > 0) {
                        sum = Double.parseDouble(middleEastAll.getIncome());
                    }
                    if (middleEastAll.getRecognizedAmount() != null && middleEastAll.getRecognizedAmount().length() > 0) {
                        income = Double.parseDouble(middleEastAll.getRecognizedAmount());
                    }
                    if (!sum.equals(income)) {
                        for (MiddleEastAll middleEastAll1 : middleEastLedgerList) {
                            if (middleEastAll.getId().equals(middleEastAll1.getId())) {
                                if (!(middleEastAll.getPi() + middleEastAll.getCustomerCode()).equals((middleEastAll1.getPi() + middleEastAll1.getCustomerCode()))) {
                                    income = income + Double.parseDouble(middleEastAll1.getRecognizedAmount());
                                }
                            }
                        }
                        if (sum.equals(income)) {
                            if ("".equals(status)) {
                                middleEastAll.setSapIncomeNo("校验成功!");
                            }
                        }
                    } else {
                        if ("".equals(middleEastAll.getSapIncomeNo())) {
                            if ("".equals(status)) {
                                middleEastAll.setSapIncomeNo("校验成功!");
                            }
                        }
                    }
                }
            }

//            ================================================校验2==========================================================================================
        }
        return middleEastLedgerList;
    }
}