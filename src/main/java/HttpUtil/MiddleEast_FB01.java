package HttpUtil;

import DataClean.BaseUtil;
import DataClean.MiddleEastAll;
import DataClean.MiddleEastSAP;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class MiddleEast_FB01 extends BaseUtil {

    public static void main(String[] args) {
//        String result = args[0];
//        String filePath = args[1];
//        String filePath_sap = args[2];
//        String filePath_cost = args[3];
//        String result = "other";
        String result = "normal";
        String filePath = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\中东AR\\test\\201907032309\\AR_Matching_List - Copy(3).xlsx";//全量台账数据表
        String filePath_sap = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\中东AR\\test\\201907032309\\FB01_2.xlsx";//生成sap表路径
        String filePath_cost = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\中东AR\\test\\20190411\\MappingList.xls";//对照表数据
        String F51Path = filePath_sap.substring(0, filePath_sap.lastIndexOf("\\") + 1) + "F-51.xlsx";
        try {
            // 获取台账数据
            List<MiddleEastAll> MiddleEastLedgerList = getLedgerExcelData(filePath);
            //生成FB01数据
            List<MiddleEastSAP> middleEastSAPList = getMiddleEastSAP(MiddleEastLedgerList, filePath, filePath_cost, result, F51Path);

            //再生成FB01数据时需要判断当前当前的时间与DocumentDateshi时间是否为同年月
            //若不是 则生成的FB01实际的PostingDate时间为本月第一天
            SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
            //获取当前年月
            Calendar calendar = Calendar.getInstance();
            calendar.setTime(new Date());
            int year = calendar.get(Calendar.YEAR);
            int month = calendar.get(Calendar.MONTH) + 1;
            for (MiddleEastSAP middleEastSAP : middleEastSAPList) {
                //获取 DocumentDate 时间的年月
                Date newDate = sdf.parse(middleEastSAP.getDocumentDate());
                calendar.setTime(newDate);
                int docYear = calendar.get(Calendar.YEAR);
                int docMonth = calendar.get(Calendar.MONTH) + 1;
                if (year != docYear || month != docMonth) {
                    calendar.setTime(new Date());
                    calendar.set(Calendar.DAY_OF_MONTH, 1);//当前日期既为本月第一天
                    middleEastSAP.setPostingDate(sdf.format(calendar.getTime()));
                }
            }
            //生成SAP表
            excelOutputSAP_AR(middleEastSAPList, filePath_sap);
            System.out.println("数据处理结束!");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 获取生成SAP数据
     *
     * @param MELedgerList
     * @param filePath
     *
     * @param filePath_Cost
     * @return
     * @throws Exception
     */
    public static List getMiddleEastSAP(List<MiddleEastAll> MELedgerList, String filePath, String filePath_Cost, String result, String F51Path) throws Exception {
        //台账中id相同的数据的处理
        Map<String, List<MiddleEastAll>> map = HandleCellData(MELedgerList);

        List<MiddleEastAll> MiddleEastListNo = new ArrayList<>();//台账数据中未认领的数据
        List<MiddleEastAll> MiddleEastTtList = new ArrayList<>();//未认领TT类型的数据集合
        List<MiddleEastAll> MiddleEastLcList = new ArrayList<>();//未认领LC类型的数据集合
        List<MiddleEastAll> MEIncomeTtList = new ArrayList<>();//TT类型数据收款不为空付款为空的所有数据
        List<MiddleEastAll> MENowIncomeTtList = new ArrayList<>();//TT类型数据收款不为空付款为空的有效数据
        List<MiddleEastAll> MENowChargeTtList = new ArrayList<>();//收款为空付款不为空的有效数据
        List<MiddleEastAll> MEOthChargeTtList = new ArrayList<>();//收款为空付款不为空的有效数据 含有OT+14位数字的数据
        List<MiddleEastAll> MEChargeTtList = new ArrayList<>();//TT类型数据收款为空付款不为空的数据
        List<MiddleEastAll> MEHistoryTtList = new ArrayList<>();//TT类型所有收款不为空的的数据
        List<MiddleEastAll> MEOtherList = new ArrayList<>();//银行附言为汇入汇款解付的数据
        List<MiddleEastAll> MEForOtherTtIncomeList = new ArrayList<>();//上月月末记为一次性客户类型的数据，本月认领之后需做处理
        List<MiddleEastAll> MEForOtherTtChargeList = new ArrayList<>();//上月月末记为一次性客户类型的数据，本月认领之后需做处理
        Set set = new HashSet();//定义set来标识单元格合并的数据
        Set set2 = new HashSet();//定义set来标识单元格合并的数据
        Set set3 = new HashSet();//定义set来标识单元格合并的数据
        //判断数据是否认领
        for (MiddleEastAll middleEastAll : MELedgerList) {
            Pattern pp = Pattern.compile("[^0-9]");//判断sapIncomeNo中是否包含10位数字编码
            String incomeNo = pp.matcher(middleEastAll.getSapIncomeNo()).replaceAll("").trim();
            //将所有sapIncomeNo中不包含10位数字编码数据的sapIncomeNo设置为空
            if ((incomeNo.length() < 10 && !"".equals(middleEastAll.getSapIncomeNo()) && !middleEastAll.getSapIncomeNo().contains("未填写StaffName"))  ) {
                middleEastAll.setSapIncomeNo("");
            }
            //过滤掉SapIncomeNo为空的数据
            if ("".equals(middleEastAll.getSapIncomeNo())) {
                String status = "";
                String summary = middleEastAll.getSummary();
                String tTLcMark = middleEastAll.getTtLCMark();
                if (!"".equals(summary) && !summary.contains("会计杂项交易录入") && !summary.contains("付款失败")) {
                    if (tTLcMark != null && !"".equals(tTLcMark) && !summary.contains("汇入汇款解付")) {
                        if ("TT".equals(tTLcMark)) {
                            MiddleEastTtList.add(middleEastAll);
                        }
                        if ("LC".equals(tTLcMark)) {
                            MiddleEastLcList.add(middleEastAll);
                        }
                    } else {
                        if (summary.contains("汇入汇款解付")) {
                            MEOtherList.add(middleEastAll);
                        } else {
                            status = status + "不属于RPA处理的业务类型";
                            middleEastAll.setSapIncomeNo(status);
                        }
                    }
                } else {
                    status = status + "不属于RPA处理的业务类型";
                    middleEastAll.setSapIncomeNo(status);
                }
                MiddleEastListNo.add(middleEastAll);
            } else {
                if (!middleEastAll.getSapIncomeNo().contains("未填写StaffName")){
                    if (middleEastAll.getIncome() != null && !"".equals(middleEastAll.getIncome()) && "TT".equals(middleEastAll.getTtLCMark())) {
                        //此处需要对状态标识字段进行一个判断，判断当前数据是否有效数据
                        Pattern p = Pattern.compile("[^0-9]");
                        Matcher m = p.matcher(middleEastAll.getSapIncomeNo());
                        if (m.matches()) {
                            //id重复的数据在历史数据集合存放一条就可
                            if (!set.contains(middleEastAll.getId())) {
                                set.add(middleEastAll.getId());
                                MEHistoryTtList.add(middleEastAll);
                            }
                        }
                    }
                }
            }
        }
        //未认领数据TT类型数据
//        System.out.println("输出tt数据条数：" + MiddleEastTtList.size());
        for (MiddleEastAll middleEastTt : MiddleEastTtList) {
            String status = "";
            //收款数据
            if (middleEastTt.getIncome() != null && !"".equals(middleEastTt.getIncome())) {
                if ("".equals(middleEastTt.getCustomerName()) || middleEastTt.getCustomerName() == null) {
                    status = status + "客户名称为空!";
                }
                if ("".equals(middleEastTt.getCustomerCode()) || middleEastTt.getCustomerCode() == null) {
                    status = status + "客户代码为空!";
                }
                if ("".equals(middleEastTt.getRecognizedAmount()) || middleEastTt.getRecognizedAmount() == null) {
                    status = status + "认领金额为空!";
                }
                if ("".equals(middleEastTt.getPi()) || middleEastTt.getPi() == null) {
                    status = status + "PI为空!";
                }
                if ("".equals(middleEastTt.getProductCode()) || middleEastTt.getProductCode() == null) {
                    status = status + "产品线为空!";
                }
                if ("".equals(middleEastTt.getPrepayment()) || middleEastTt.getPrepayment() == null) {
                    status = status + "预付比例为空!";
                } else {
                    if (!middleEastTt.getPrepayment().contains("N") && !middleEastTt.getPrepayment().contains("Y")
                            && !middleEastTt.getPrepayment().contains("尾款")) {
                        status = status + "预付比例不符合规则!";
                    }
                }
                //判断数据是有效数据
                if ("".equals(status)) {
                    //本月认领的一次性客户类型数据
                    if (!"".equals(middleEastTt.getSapNoForOther()) && middleEastTt.getSapNoForOther() != null) {
                        List<MiddleEastAll> lists = map.get(middleEastTt.getId());
                        DecimalFormat df = new DecimalFormat("#.00");
                        if (lists != null && lists.size() > 0) {
                            double sum = 0;
                            System.out.println(lists.size());
                            for (MiddleEastAll middleEast : lists) {
                                if (!set3.contains(middleEast.getCustomerCode())) {
                                    set3.add(middleEast.getCustomerCode());
                                }
                                if (middleEast.getRecognizedAmount() != null && !"".equals(middleEast.getRecognizedAmount())) {

                                    sum = sum + Double.parseDouble(middleEast.getRecognizedAmount());
                                }
                            }
                            System.out.println("sum:" + sum);
                            if (set3.size() == 1) {
                                sum = Double.parseDouble(df.format(sum));
                                Double income = Double.parseDouble(middleEastTt.getIncome());
                                if (sum != income) {
                                    status = status + "认领金额与收款金额不一致";
                                    middleEastTt.setSapIncomeNo(status);
                                } else {
                                    //单元格合并的数据在历史数据集合存放一条就可
                                    if (!set.contains(middleEastTt.getId())) {
                                        set.add(middleEastTt.getId());
                                        MEHistoryTtList.add(middleEastTt);
                                    }
                                }
                            } else {
                                status = status + "客户代码不一致";
                                middleEastTt.setSapIncomeNo(status);
                            }
                        } else {
                            double income = Double.parseDouble(middleEastTt.getIncome());
                            double recognizedAmount = Double.parseDouble(middleEastTt.getRecognizedAmount());
                            if (income == recognizedAmount) {
                                MEHistoryTtList.add(middleEastTt);
                                //未处理数据
                                MEForOtherTtIncomeList.add(middleEastTt);//一次性客户认领的数据
                            } else {
                                status = status + "认领金额与收款金额不一致";
                                middleEastTt.setSapIncomeNo(status);
                            }
                        }
                    } else {
                        //若数据是有效数据，通过数据的唯一标识ID，查看该数据是否是单元格合并数据（）
                        List<MiddleEastAll> lists = map.get(middleEastTt.getId());
                        DecimalFormat df = new DecimalFormat("#.00");
                        if (lists != null && lists.size() > 0) {
                            double sum = 0;
                            for (MiddleEastAll middleEast : lists) {
                                if (!set2.contains(middleEast.getCustomerCode())) {
                                    set2.add(middleEast.getCustomerCode());
                                }
                                if (middleEast.getRecognizedAmount() != null && !"".equals(middleEast.getRecognizedAmount())) {
                                    sum = sum + Double.parseDouble(middleEast.getRecognizedAmount());
                                }
                            }
                            if (set2.size() == 1) {
                                sum = Double.parseDouble(df.format(sum));
                                Double income = Double.parseDouble(middleEastTt.getIncome());
                                if (sum == income) {
                                    //单元格合并的数据在历史数据集合存放一条就可
                                    if (!set.contains(middleEastTt.getId())) {
                                        set.add(middleEastTt.getId());
                                        MEHistoryTtList.add(middleEastTt);
                                    }
                                } else {
                                    status = status + "认领金额与收款金额不一致";
                                    middleEastTt.setSapIncomeNo(status);
                                }
                            } else {
                                status = status + "客户代码不一致";
                                middleEastTt.setSapIncomeNo(status);
                            }
                        } else {
                            double income = Double.parseDouble(middleEastTt.getIncome());
                            double recognizedAmount = Double.parseDouble(middleEastTt.getRecognizedAmount());
                            if (income == recognizedAmount) {
                                MEHistoryTtList.add(middleEastTt);
                                //未处理数据
                                MENowIncomeTtList.add(middleEastTt);
                            } else {
                                status = status + "认领金额与收款金额不一致";
                                middleEastTt.setSapIncomeNo(status);
                            }
                        }
                    }
                } else {
                    middleEastTt.setSapIncomeNo(status);
                }
                MEIncomeTtList.add(middleEastTt);

            }
            //付款数据
            if (middleEastTt.getCharge() != null && !"".equals(middleEastTt.getCharge())) {
                middleEastTt.setSapIncomeNo("未找到正向收款金额或正向收款金额不规范");
                MEChargeTtList.add(middleEastTt);
            }
        }


        //通过收款金额是负数的筛选逻辑，获取所有的有效数据
        for (MiddleEastAll middleEastF : MEChargeTtList) {
            String str = "";
            Pattern p = Pattern.compile("(IT[0-9]{14})");
            Pattern p2 = Pattern.compile("(TI[0-9]{14})");
            Pattern p3 = Pattern.compile("(BACDUB[0-9]{6})");
            Pattern p4 = Pattern.compile("(OT[0-9]{14})");
            if (!"".equals(getData(middleEastF.getSummary(), p))) {
                str = getData(middleEastF.getSummary(), p);
            } else if (!"".equals(getData(middleEastF.getSummary(), p2))) {
                str = getData(middleEastF.getSummary(), p2);
            } else {
                str = getData(middleEastF.getSummary(), p3);
            }
            if (!"".equals(str)) {
                for (MiddleEastAll middleEastZ : MEHistoryTtList) {
                    System.out.println("middleEastZ.getSummary()"+middleEastZ.getSummary());
                    if (middleEastZ.getSummary().contains(str)) {
                        middleEastF.setCustomerName(middleEastZ.getCustomerName());
                        middleEastF.setCustomerCode(middleEastZ.getCustomerCode());
                        middleEastF.setProductCode(middleEastZ.getProductCode());
                        middleEastF.setPi(middleEastZ.getPi());
                        middleEastF.setSapIncomeNo("");
                        if (middleEastF.getSapNoForOther() != null && !"".equals(middleEastF.getSapNoForOther())) {
                            if (!("14".equals(middleEastF.getSapNoForOther().substring(0, 2)) && (middleEastF.getSapClearingNo() == null || "".equals(middleEastF.getSapClearingNo())))) {
                                MEForOtherTtChargeList.add(middleEastF);
                            }
//                                MEForOtherTtChargeList.add(middleEastF);

                        } else {
                            MENowChargeTtList.add(middleEastF);
                        }
                    }
                }
            } else if (!"".equals(getData(middleEastF.getSummary(), p4))) {
                MEOthChargeTtList.add(middleEastF);
                middleEastF.setSapIncomeNo("");
            }
        }
        //TT类型数据修改状态
        for (MiddleEastAll middleEast : MiddleEastTtList) {
            for (MiddleEastAll middleEast1 : MEChargeTtList) {
                if (middleEast.getId().equals(middleEast1.getId())) {
                    middleEast.setSapIncomeNo(middleEast1.getSapIncomeNo());
                }
            }
        }
        for (MiddleEastAll middleEast : MiddleEastListNo) {//未认领数据状态修改
            for (MiddleEastAll middleEastTt : MiddleEastTtList) {
                if (middleEast.getId().equals(middleEastTt.getId())) {
                    middleEast.setSapIncomeNo(middleEastTt.getSapIncomeNo());
                }
            }
        }
        //获取TT类型数据的SAP数据
        List<MiddleEastSAP> middleEastTtSAPList = cleanTtDataSap(MENowIncomeTtList, MEIncomeTtList, MENowChargeTtList, MEOthChargeTtList, map, filePath_Cost).get("List");
        List<MiddleEastAll> MiddleEastListTT = cleanTtDataSap(MENowIncomeTtList, MEIncomeTtList, MENowChargeTtList, MEOthChargeTtList, map, filePath_Cost).get("ledger");
//        List<MiddleEastSAP> middleEastF51TtSAPList = cleanTtDataSap(MENowIncomeTtList,MEIncomeTtList, MENowChargeTtList, MEOthChargeTtList, map, filePath_Cost).get("middleEastSAP_F_51_List");
        //一次性客户类型数据本月认领
//        System.out.println("一次性客户：" +MEForOtherTtIncomeList.size());
        System.out.println(MEForOtherTtIncomeList.size());
        Map<String, List> map1 = cleanForOtherTtDataSap(MEForOtherTtIncomeList, MEIncomeTtList, MEForOtherTtChargeList, map, filePath_Cost);
        List<MiddleEastSAP> middleEastForOtherTtSAPList = map1.get("middleEastSAPList");
        List<MiddleEastSAP> middleEastF51TtSAPList = map1.get("middleEastSAP_F_51_List");
        middleEastTtSAPList.addAll(middleEastForOtherTtSAPList);


        //获取Lc类型数据的SAP
        List<MiddleEastSAP> middleEastLcSAPList = new ArrayList<>();
        for (MiddleEastAll middleEC : MiddleEastLcList) {
            List<MiddleEastAll> lists = map.get(middleEC.getId());
            if (lists == null || lists.size() == 0) {
                MiddleEastSAP middleEastSAP = new MiddleEastSAP();
                MiddleEastSAP middleEastSAP1 = new MiddleEastSAP();
                String charge = middleEC.getCharge();
                String income = middleEC.getIncome();
                String customerCode = middleEC.getCustomerCode();
                String lCInvoiceNumber = middleEC.getSummary();
                String status = "";
                //扣款金额为空，收款金额不为空
                if ((charge == null || "".equals(charge)) && (income != null && !"".equals(income))) {
                    if (customerCode != null && !"".equals(customerCode)) {
                        if (customerCode.contains("6000") || customerCode.contains("A05") || customerCode.contains("A06") || customerCode.contains("A07")) {
                            String uuid = UUID.randomUUID().toString().replaceAll("-", "");
                            middleEastSAP.setNo(uuid);
                            middleEastSAP.setId(middleEC.getId());
                            middleEastSAP.setDocumentDate(middleEC.getDocumentDate());
                            middleEastSAP.setType("DZ");
                            middleEastSAP.setCompanyCode("6770");
                            middleEastSAP.setPostingDate(middleEC.getDocumentDate());
                            int i = middleEC.getDocumentDate().indexOf("/", 0);
                            int j = middleEC.getDocumentDate().lastIndexOf("/");
                            String date = middleEC.getDocumentDate();
                            SimpleDateFormat sd = new SimpleDateFormat("MM/dd/yyy");
                            SimpleDateFormat sdfd = new SimpleDateFormat("dd/MM");
                            date = sdfd.format(sd.parse(date));
                            middleEastSAP.setPeriod(middleEC.getDocumentDate().substring(0, i));
                            middleEastSAP.setCurrencyRate(middleEC.getCurrency());
                            String gl = middleEC.getBank().substring(middleEC.getBank().indexOf(" ") + 1);

                            middleEastSAP1.setNo(uuid);
                            middleEastSAP1.setId(middleEC.getId());
                            middleEastSAP1.setDocumentDate(middleEC.getDocumentDate());
                            middleEastSAP1.setType("DZ");
                            middleEastSAP1.setCompanyCode("6770");
                            middleEastSAP1.setPostingDate(middleEC.getDocumentDate());
                            middleEastSAP1.setPeriod(middleEC.getDocumentDate().substring(0, i));
                            middleEastSAP1.setCurrencyRate(middleEC.getCurrency());

                            if (lCInvoiceNumber.contains("123") && lCInvoiceNumber.indexOf("123", 0) == 0 && lCInvoiceNumber.contains(" ")) {
                                String[] vl = lCInvoiceNumber.split(" ");
                                List<String> list = Arrays.asList(vl);
                                middleEastSAP.setAssignment(list.get(list.size() - 1));
                                middleEastSAP.setText(list.get(list.size() - 1));
                                middleEastSAP.setLongText(list.get(list.size() - 1));

                                middleEastSAP1.setAssignment(list.get(list.size() - 1));
                                middleEastSAP1.setText(list.get(list.size() - 1));
                                middleEastSAP1.setLongText(list.get(list.size() - 1));
                            } else {
                                middleEastSAP.setAssignment(lCInvoiceNumber);
                                middleEastSAP.setText(lCInvoiceNumber);
                                middleEastSAP.setLongText(lCInvoiceNumber);

                                middleEastSAP1.setAssignment(lCInvoiceNumber);
                                middleEastSAP1.setText(lCInvoiceNumber);
                                middleEastSAP1.setLongText(lCInvoiceNumber);
                            }
                            if ((customerCode.equals("A01")) && (middleEC.getSapNoForOther().equals("")|| middleEC.getSapNoForOther()!=null)){
                                status = middleEC.getCustomerCode();
                            }
                            if (customerCode.contains("6000") && customerCode.indexOf("6000", 0) == 0 && customerCode.length() == 10) {
                                if (middleEC.getCustomerName() != null && !"".equals(middleEC.getCustomerName())) {
                                    //一次性客户类型数据
                                    lCInvoiceNumber = middleEC.getSummary();
                                    String[] vl = lCInvoiceNumber.split(" ");
                                    List<String> list = Arrays.asList(vl);
                                    if (middleEC.getSapNoForOther() != null && !"".equals(middleEC.getSapNoForOther())) {
                                        DateFormat dFormat = new SimpleDateFormat("MM/dd/yyyy");
                                        Calendar calendar = Calendar.getInstance();
                                        calendar.setTime(new Date());
                                        calendar.set(Calendar.DAY_OF_MONTH, 1);//当前日期既为本月第一天
                                        String month = String.valueOf(calendar.get(Calendar.MONTH) + 1);
                                        SimpleDateFormat sd1 = new SimpleDateFormat("dd/MM");
                                        String date1 = sd1.format(calendar.getTime());
                                        middleEastSAP.setPostingDate(dFormat.format(calendar.getTime()));
                                        middleEastSAP.setPeriod(month);
                                        middleEastSAP.setReference(middleEC.getCustomerName());
                                        middleEastSAP.setDocHeaderText(middleEC.getCustomerName() + " " + date1);
                                        middleEastSAP.setPostKey("09");
                                        middleEastSAP.setGl("6000007049");
                                        middleEastSAP.setSglInd(">");
                                        middleEastSAP.setAmount(convertData(middleEC.getIncome()));
                                        middleEastSAP.setTaxCode("X0");
                                        middleEastSAP.setDuOn(middleEC.getDocumentDate());
                                        middleEastSAP.setReasonCode("112");

                                        middleEastSAP1.setPostingDate(dFormat.format(calendar.getTime()));
                                        middleEastSAP1.setPeriod(month);
                                        middleEastSAP1.setReference(middleEC.getCustomerName());
                                        middleEastSAP1.setDocHeaderText(middleEC.getCustomerName() + " " + date1);
                                        middleEastSAP1.setPostKey("15");

                                        middleEastSAP1.setAssignment(list.get(list.size() - 1));
                                        middleEastSAP.setText(list.get(list.size() - 1));
                                        middleEastSAP.setLongText(list.get(list.size() - 1));

                                        middleEastSAP1.setGl(middleEC.getCustomerCode());
                                        middleEastSAP1.setSglInd("");
                                        middleEastSAP1.setAmount(convertData(middleEC.getIncome()));
                                        middleEastSAP1.setTaxCode("");
                                        middleEastSAP1.setSapforother(middleEC.getSapNoForOther());
                                        middleEastSAP1.setBlinDate(middleEC.getDocumentDate());
                                        if (!(middleEC.getSapNoForOther().substring(0, 2).equals("14") && (middleEC.getSapClearingNo() == null || middleEC.getSapClearingNo().equals("")) && !("A01".equals(middleEC.getCustomerCode())))) {
                                            middleEastLcSAPList.add(middleEastSAP);
                                            middleEastLcSAPList.add(middleEastSAP1);
                                        } else {
                                            middleEastF51TtSAPList.add(middleEastSAP1);
                                        }
                                    } else {
                                        middleEastSAP.setReference(middleEC.getCustomerName());
                                        middleEastSAP.setDocHeaderText(middleEC.getCustomerName() + " " + date);
                                        middleEastSAP.setPostKey("40");
                                        middleEastSAP.setAmount(convertData(middleEC.getIncome()));
                                        middleEastSAP.setGl(gl);
                                        middleEastSAP.setValueDate(middleEC.getDocumentDate());
                                        middleEastSAP.setReasonCode("112");

                                        middleEastSAP1.setReference(middleEC.getCustomerName());
                                        middleEastSAP1.setDocHeaderText(middleEC.getCustomerName() + " " + date);
                                        middleEastSAP1.setPostKey("15");
                                        middleEastSAP1.setGl(middleEC.getCustomerCode());
                                        middleEastSAP1.setAmount(convertData(middleEC.getIncome()));
                                        middleEastSAP1.setBlinDate(middleEC.getDocumentDate());

                                        middleEastLcSAPList.add(middleEastSAP);
                                        middleEastLcSAPList.add(middleEastSAP1);
                                    }
                                } else {
                                    status = status + "客户名称为空!";
                                }
                            }
                            if ("A05".equals(customerCode)) {
                                if (middleEC.getDepositPrincipal() != null && !"".equals(middleEC.getDepositPrincipal())) {
                                    middleEastSAP.setType("SA");
                                    middleEastSAP.setReference("Deposit" + " " + date);
                                    middleEastSAP.setDocHeaderText("Deposit" + " " + date);
                                    middleEastSAP.setPostKey("40");
                                    middleEastSAP.setGl(gl);
                                    middleEastSAP.setAmount(convertData(middleEC.getIncome()));
                                    middleEastSAP.setValueDate(middleEC.getDocumentDate());
                                    middleEastSAP.setAssignment("Deposit" + " " + date);
                                    middleEastSAP.setText("Deposit" + " " + date);
                                    middleEastSAP.setLongText("");
                                    middleEastSAP.setReasonCode("362");

                                    middleEastSAP1.setType("SA");
                                    middleEastSAP1.setReference("Deposit" + " " + date);
                                    middleEastSAP1.setDocHeaderText("Deposit" + " " + date);
                                    middleEastSAP1.setGl("1002999998");
                                    middleEastSAP1.setPostKey("50");
                                    middleEastSAP1.setAmount(convertData(middleEC.getIncome()));
                                    middleEastSAP1.setValueDate(middleEC.getDocumentDate());
                                    middleEastSAP1.setAssignment("Deposit" + " " + date);
                                    middleEastSAP1.setText("Deposit" + " " + date);
                                    middleEastSAP1.setLongText("");
                                    middleEastSAP1.setReasonCode("362");

                                    middleEastLcSAPList.add(middleEastSAP);
                                    middleEastLcSAPList.add(middleEastSAP1);

                                } else {
                                    status = status + "定存本金为空!";
                                }
                            }
                            if ("A06".equals(customerCode)) {
                                if (middleEC.getDepositInterest() != null && !"".equals(middleEC.getDepositInterest())) {
                                    middleEastSAP.setType("SA");
                                    middleEastSAP.setReference("Interest income");
                                    middleEastSAP.setDocHeaderText("Interest income");
                                    middleEastSAP.setPostKey("40");
                                    middleEastSAP.setGl(gl);
                                    middleEastSAP.setAmount(convertData(middleEC.getIncome()));
                                    middleEastSAP.setAssignment("Interest income");
                                    middleEastSAP.setText("Interest income");
                                    middleEastSAP.setLongText("");
                                    middleEastSAP.setReasonCode("332");

                                    middleEastSAP1.setType("SA");
                                    middleEastSAP1.setReference("Interest income");
                                    middleEastSAP1.setDocHeaderText("Interest income");
                                    middleEastSAP1.setGl("5503010000");
                                    middleEastSAP1.setPostKey("50");
                                    middleEastSAP1.setAmount(convertData(middleEC.getIncome()));
                                    middleEastSAP1.setCostCenter("6677110102");
                                    middleEastSAP1.setAssignment("Interest income");
                                    middleEastSAP1.setText("Interest income");
                                    middleEastSAP1.setLongText("");

                                    middleEastLcSAPList.add(middleEastSAP);
                                    middleEastLcSAPList.add(middleEastSAP1);
                                } else {
                                    status = status + "定存利息为空!";
                                }
                            }
                            if ("A07".equals(customerCode)) {
                                if ((middleEC.getDepositPrincipal() != null && !"".equals(middleEC.getDepositPrincipal()))
                                        && (middleEC.getDepositInterest() != null && !"".equals(middleEC.getDepositInterest()))) {
                                    //本金
                                    MiddleEastSAP middleEastSAP_B = new MiddleEastSAP();
                                    MiddleEastSAP middleEastSAP_B1 = new MiddleEastSAP();
                                    String uuid1 = UUID.randomUUID().toString().replaceAll("-", "");
                                    middleEastSAP_B.setNo(uuid1);
                                    middleEastSAP_B.setId(middleEC.getId());
                                    middleEastSAP_B.setDocumentDate(middleEC.getDocumentDate());
                                    middleEastSAP_B.setType("SA");
                                    middleEastSAP_B.setCompanyCode("6770");
                                    middleEastSAP_B.setPostingDate(middleEC.getDocumentDate());
                                    middleEastSAP_B.setPeriod(middleEC.getDocumentDate().substring(0, i));
                                    middleEastSAP_B.setCurrencyRate(middleEC.getCurrency());
                                    middleEastSAP_B.setReference("Deposit" + " " + date);
                                    middleEastSAP_B.setDocHeaderText("Deposit" + " " + date);
                                    middleEastSAP_B.setPostKey("40");
                                    middleEastSAP_B.setGl(gl);
                                    middleEastSAP_B.setAmount(convertData(middleEC.getDepositPrincipal()));
                                    middleEastSAP_B.setValueDate(middleEC.getDocumentDate());
                                    middleEastSAP_B.setAssignment("Deposit" + " " + date);
                                    middleEastSAP_B.setText("Deposit" + " " + date);
                                    middleEastSAP_B.setReasonCode("362");

                                    middleEastSAP_B1.setNo(uuid1);
                                    middleEastSAP_B1.setId(middleEC.getId());
                                    middleEastSAP_B1.setDocumentDate(middleEC.getDocumentDate());
                                    middleEastSAP_B1.setType("SA");
                                    middleEastSAP_B1.setCompanyCode("6770");
                                    middleEastSAP_B1.setPostingDate(middleEC.getDocumentDate());
                                    middleEastSAP_B1.setPeriod(middleEC.getDocumentDate().substring(0, i));
                                    middleEastSAP_B1.setCurrencyRate(middleEC.getCurrency());
                                    middleEastSAP_B1.setReference("Deposit" + " " + date);
                                    middleEastSAP_B1.setDocHeaderText("Deposit" + " " + date);
                                    middleEastSAP_B1.setGl("1002999998");
                                    middleEastSAP_B1.setPostKey("50");
                                    middleEastSAP_B1.setAmount(convertData(middleEC.getDepositInterest()));
                                    middleEastSAP_B1.setValueDate(middleEC.getDocumentDate());
                                    middleEastSAP_B1.setAssignment("Deposit" + " " + date);
                                    middleEastSAP_B1.setText("Deposit" + " " + date);
                                    middleEastSAP_B1.setReasonCode("362");

                                    middleEastLcSAPList.add(middleEastSAP_B);
                                    middleEastLcSAPList.add(middleEastSAP_B1);

                                    //利息
                                    MiddleEastSAP middleEastSAP_L = new MiddleEastSAP();
                                    MiddleEastSAP middleEastSAP_L1 = new MiddleEastSAP();
                                    String uuid2 = UUID.randomUUID().toString().replaceAll("-", "");
                                    middleEastSAP_L.setNo(uuid2);
                                    middleEastSAP_L.setId(middleEC.getId());
                                    middleEastSAP_L.setDocumentDate(middleEC.getDocumentDate());
                                    middleEastSAP_L.setType("SA");
                                    middleEastSAP_L.setCompanyCode("6770");
                                    middleEastSAP_L.setPostingDate(middleEC.getDocumentDate());
                                    middleEastSAP_L.setPeriod(middleEC.getDocumentDate().substring(0, i));
                                    middleEastSAP_L.setCurrencyRate(middleEC.getCurrency());
                                    middleEastSAP_L.setReference("Interest income");
                                    middleEastSAP_L.setDocHeaderText("Interest income");
                                    middleEastSAP_L.setPostKey("40");
                                    middleEastSAP_L.setGl(gl);
                                    middleEastSAP_L.setAmount(convertData(middleEC.getIncome()));
                                    middleEastSAP_L.setValueDate(middleEC.getDocumentDate());
                                    middleEastSAP_L.setAssignment("Interest income");
                                    middleEastSAP_L.setText("Interest income");
                                    middleEastSAP_L.setReasonCode("332");

                                    middleEastSAP_L1.setNo(uuid2);
                                    middleEastSAP_L1.setId(middleEC.getId());
                                    middleEastSAP_L1.setDocumentDate(middleEC.getDocumentDate());
                                    middleEastSAP_L1.setType("SA");
                                    middleEastSAP_L1.setCompanyCode("6770");
                                    middleEastSAP_L1.setPostingDate(middleEC.getDocumentDate());
                                    middleEastSAP_L1.setPeriod(middleEC.getDocumentDate().substring(0, i));
                                    middleEastSAP_L1.setCurrencyRate(middleEC.getCurrency());
                                    middleEastSAP_L1.setReference("Interest income");
                                    middleEastSAP_L1.setDocHeaderText("Interest income");
                                    middleEastSAP_L1.setPostKey("50");
                                    middleEastSAP_L1.setGl("5503010000");
                                    middleEastSAP_L1.setAmount(convertData(middleEC.getIncome()));
                                    middleEastSAP_L1.setCostCenter("6677110102");
                                    middleEastSAP_L1.setAssignment("Interest income");
                                    middleEastSAP_L1.setText("Interest income");

                                    middleEastLcSAPList.add(middleEastSAP_L);
                                    middleEastLcSAPList.add(middleEastSAP_L1);
                                } else {
                                    if (middleEC.getDepositPrincipal() == null || "".equals(middleEC.getDepositPrincipal())) {
                                        status = status + "定存本金为空!";
                                    }
                                    if (middleEC.getDepositInterest() == null || "".equals(middleEC.getDepositInterest())) {
                                        status = status + "定存利息为空!";
                                    }
                                }
                            }
                        } else {
                            status = status + "未找到相应业务类型!";
                        }
                    } else {
                        status = status + "客户代码为空!";
                    }
                }
                //扣款金额不为空，收款金额为空
                if ((charge != null && !"".equals(charge)) && (income == null || "".equals(income))) {
                    if (customerCode != null && !"".equals(customerCode)) {
                        if (customerCode.contains("6000") || customerCode.contains("A01") || customerCode.contains("A02")
                                || customerCode.contains("A03") || customerCode.contains("A04") || customerCode.contains("A08")) {
                            if (lCInvoiceNumber.contains("123") && lCInvoiceNumber.indexOf("123", 0) == 0 && lCInvoiceNumber.contains(" ")) {
                                String[] vl = lCInvoiceNumber.split(" ");
                                List<String> list = Arrays.asList(vl);
                                middleEastSAP.setAssignment(list.get(list.size() - 1));
                                middleEastSAP.setLongText(list.get(list.size() - 1));
                                middleEastSAP.setText(list.get(list.size() - 1));

                                middleEastSAP1.setAssignment(list.get(list.size() - 1));
                                middleEastSAP1.setText(list.get(list.size() - 1));
                                middleEastSAP1.setLongText(list.get(list.size() - 1));
                            } else {
                                middleEastSAP.setAssignment(lCInvoiceNumber);
                                middleEastSAP.setLongText(lCInvoiceNumber);
                                middleEastSAP.setText(lCInvoiceNumber);

                                middleEastSAP1.setAssignment(lCInvoiceNumber);
                                middleEastSAP1.setText(lCInvoiceNumber);
                                middleEastSAP1.setLongText(lCInvoiceNumber);
                            }

                            String uuid = UUID.randomUUID().toString().replaceAll("-", "");
                            middleEastSAP.setNo(uuid);
                            middleEastSAP.setId(middleEC.getId());
                            middleEastSAP.setDocumentDate(middleEC.getDocumentDate());
                            middleEastSAP.setCompanyCode("6770");
                            middleEastSAP.setPostingDate(middleEC.getDocumentDate());
                            int i = middleEC.getDocumentDate().indexOf("/", 0);
                            int j = middleEC.getDocumentDate().lastIndexOf("/");
                            String date = middleEC.getDocumentDate();
                            SimpleDateFormat sd = new SimpleDateFormat("MM/dd/yyy");
                            SimpleDateFormat sdfd = new SimpleDateFormat("dd/MM");
                            date = sdfd.format(sd.parse(date));
                            middleEastSAP.setPeriod(middleEC.getDocumentDate().substring(0, i));
                            middleEastSAP.setCurrencyRate(middleEC.getCurrency());
                            String gl = middleEC.getBank().substring(middleEC.getBank().indexOf(" ") + 1);

                            middleEastSAP1.setId(middleEC.getId());
                            middleEastSAP1.setNo(uuid);
                            middleEastSAP1.setDocumentDate(middleEC.getDocumentDate());
                            middleEastSAP1.setCompanyCode("6770");
                            middleEastSAP1.setPostingDate(middleEC.getDocumentDate());
                            middleEastSAP1.setPeriod(middleEC.getDocumentDate().substring(0, i));
                            middleEastSAP1.setCurrencyRate(middleEC.getCurrency());
                            //客户代码为6000开头的10位编码
                            if (customerCode.contains("6000") && customerCode.indexOf("6000", 0) == 0 && customerCode.length() == 10) {
                                if (middleEC.getCustomerName() != null && !"".equals(middleEC.getCustomerName())) {
                                    if (middleEC.getSapNoForOther() != null && !"".equals(middleEC.getSapNoForOther())) {
                                        DateFormat dFormat = new SimpleDateFormat("MM/dd/yyyy");
                                        Calendar calendar = Calendar.getInstance();
                                        calendar.setTime(new Date());
                                        calendar.set(Calendar.DAY_OF_MONTH, 1);//当前日期既为本月第一天
                                        String month = String.valueOf(calendar.get(Calendar.MONTH) + 1);
                                        SimpleDateFormat sd1 = new SimpleDateFormat("dd/MM");
                                        String date1 = sd1.format(calendar.getTime());
                                        middleEastSAP.setType("SA");
                                        middleEastSAP.setPostingDate(dFormat.format(calendar.getTime()));
                                        middleEastSAP.setPeriod(month);
                                        middleEastSAP.setReference(middleEC.getCustomerCode());
                                        middleEastSAP.setDocHeaderText("BANK CHARGE" + " " + date1);
                                        middleEastSAP.setPostKey("50");
                                        middleEastSAP.setGl("5503040000");
                                        middleEastSAP.setAmount(convertData(middleEC.getCharge()));
                                        middleEastSAP.setCostCenter("6677110102");

                                        middleEastSAP1.setPostingDate(dFormat.format(calendar.getTime()));
                                        middleEastSAP1.setPeriod(month);
                                        middleEastSAP1.setReference(middleEC.getCustomerCode());
                                        middleEastSAP1.setDocHeaderText("BANK CHARGE" + " " + date1);
                                        middleEastSAP1.setPostKey("40");
                                        middleEastSAP1.setGl("5503040000");
                                        middleEastSAP1.setSglInd("");
                                        middleEastSAP1.setAmount(convertData(middleEC.getCharge()));
                                        middleEastSAP1.setCostCenter(getCostCenter(middleEC.getProductCode(), filePath_Cost));

                                        if ("".equals(middleEastSAP1.getCostCenter()) && !"".equals(middleEC.getProductCode())) {
                                            status = status + "ProductCode填写有误,CostCenter无法正常匹配!";
                                        } else {
                                            middleEastLcSAPList.add(middleEastSAP1);
                                            middleEastLcSAPList.add(middleEastSAP);
                                        }


                                    } else {
                                        middleEastSAP.setType("SA");
                                        middleEastSAP.setDocHeaderText("BANK CHARGE" + " " + date);
                                        middleEastSAP.setPostKey("50");
                                        middleEastSAP.setReference(middleEC.getCustomerCode());
                                        middleEastSAP.setGl(gl);
                                        middleEastSAP.setAmount(convertData(middleEC.getCharge()));
                                        middleEastSAP.setValueDate(middleEC.getDocumentDate());
                                        middleEastSAP.setReasonCode("179");

                                        middleEastSAP1.setType("SA");
                                        middleEastSAP1.setPostKey("40");
                                        middleEastSAP1.setDocHeaderText("BANK CHARGE" + " " + date);
                                        middleEastSAP1.setReference(middleEC.getCustomerCode());
                                        middleEastSAP1.setCostCenter(getCostCenter(middleEC.getProductCode(), filePath_Cost));
                                        middleEastSAP1.setGl("5503040000");
                                        middleEastSAP1.setAmount(convertData(middleEC.getCharge()));

                                        if ("".equals(middleEastSAP1.getCostCenter()) && !"".equals(middleEC.getProductCode())) {
                                            status = status + "ProductCode填写有误,CostCenter无法正常匹配!";
                                        } else {
                                            middleEastLcSAPList.add(middleEastSAP1);
                                            middleEastLcSAPList.add(middleEastSAP);
                                        }
                                    }
                                } else {
                                    status = status + "客户名称为空!";
                                }
                            }
                            if (customerCode.contains("A04")) {
                                if (middleEC.getTransferTo() != null && !"".equals(middleEC.getTransferTo())) {
                                    middleEastSAP.setType("SA");
                                    middleEastSAP.setPostKey("40");
                                    middleEastSAP.setReference("Transfer");
                                    middleEastSAP.setDocHeaderText("Transfer");
                                    middleEastSAP.setAmount(convertData(middleEC.getCharge()));
                                    middleEastSAP.setGl(getAccountDetails(middleEC.getTransferTo(), filePath_Cost));
                                    String s = middleEC.getBank().substring(0, middleEC.getBank().indexOf(" "));
                                    s = "From " + s + " To " + middleEC.getTransferTo();
                                    middleEastSAP.setAssignment(s);
                                    middleEastSAP.setValueDate(middleEC.getDocumentDate());
                                    middleEastSAP.setText(s);
                                    middleEastSAP.setLongText("");
                                    middleEastSAP.setReasonCode("100");

                                    middleEastSAP1.setType("SA");
                                    middleEastSAP1.setReference("Transfer");
                                    middleEastSAP1.setDocHeaderText("Transfer");
                                    middleEastSAP1.setPostKey("50");
                                    middleEastSAP1.setAmount(convertData(middleEC.getCharge()));
                                    middleEastSAP1.setValueDate(middleEC.getDocumentDate());
                                    middleEastSAP1.setGl(gl);
                                    middleEastSAP1.setAssignment(s);
                                    middleEastSAP1.setText(s);
                                    middleEastSAP1.setLongText("");
                                    middleEastSAP1.setReasonCode("100");

                                    middleEastLcSAPList.add(middleEastSAP);
                                    middleEastLcSAPList.add(middleEastSAP1);
                                } else {
                                    status = status + "内部转账为空!";
                                }
                            }
                            //客户代码为“Common”或者“手续费”
//                            if ((customerCode.equals("A01")) && (middleEC.getSapNoForOther().equals("")|| middleEC.getSapNoForOther()!=null)){
//                                status = middleEC.getSapNoForOther();
//                                middleEastLcSAPList.add(middleEastSAP);
//                                middleEastLcSAPList.add(middleEastSAP1);
//                            }
                            if ((customerCode.equals("A01")) && (middleEC.getSapNoForOther().equals("")|| middleEC.getSapNoForOther()!=null)){
                                middleEastSAP.setType("SA");
                                middleEastSAP.setPostKey("50");
                                middleEastSAP.setReference("BANK CHARGE");
                                middleEastSAP.setDocHeaderText("BANK CHARGE" + " " + date);
                                middleEastSAP.setGl(gl);
                                middleEastSAP.setAmount(convertData(middleEC.getCharge()));
                                middleEastSAP.setValueDate(middleEC.getDocumentDate());
                                middleEastSAP.setReasonCode("179");

                                middleEastSAP1.setType("SA");
                                middleEastSAP1.setReference("BANK CHARGE");
                                middleEastSAP1.setDocHeaderText("BANK CHARGE" + " " + date);
                                middleEastSAP1.setCostCenter("6677110102");
                                middleEastSAP1.setPostKey("40");
                                middleEastSAP1.setGl("5503040000");
                                middleEastSAP1.setAmount(convertData(middleEC.getCharge()));

                                middleEastLcSAPList.add(middleEastSAP);
                                middleEastLcSAPList.add(middleEastSAP1);
                            }
                            //信用卡还款
                            if ("A02".equals(customerCode)) {
                                middleEastSAP.setType("KZ");
                                middleEastSAP.setPostKey("50");
                                middleEastSAP.setReference("HSBC Card");
                                middleEastSAP.setDocHeaderText("HSBC Card");
                                middleEastSAP.setAmount(convertData(middleEC.getCharge()));
                                middleEastSAP.setGl(gl);
                                middleEastSAP.setValueDate(middleEC.getDocumentDate());
                                middleEastSAP.setReasonCode("179");

                                middleEastSAP1.setType("KZ");
                                middleEastSAP1.setReference("HSBC Card");
                                middleEastSAP1.setDocHeaderText("HSBC Card");
                                middleEastSAP1.setPostKey("25");
                                middleEastSAP1.setGl("V999990906");
                                middleEastSAP1.setAmount(convertData(middleEC.getCharge()));
                                middleEastSAP1.setBlinDate(middleEC.getDocumentDate());

                                middleEastLcSAPList.add(middleEastSAP);
                                middleEastLcSAPList.add(middleEastSAP1);
                            }
                            //工资类型
                            if ("A03".equals(customerCode)) {
                                BigDecimal bg = new BigDecimal(charge);
                                double charge1 = bg.setScale(2, BigDecimal.ROUND_HALF_UP).doubleValue();
//                                if (charge1 < 60000){
                                MiddleEastSAP middleEastSAP_S = new MiddleEastSAP();
                                MiddleEastSAP middleEastSAP_S1 = new MiddleEastSAP();
                                String uuids = UUID.randomUUID().toString().replaceAll("-", "");
                                middleEastSAP_S.setId(middleEC.getId());
                                middleEastSAP_S.setNo(uuids);
                                middleEastSAP_S.setDocumentDate(middleEC.getDocumentDate());
                                middleEastSAP_S.setCompanyCode("6770");
                                middleEastSAP_S.setType("SA");
                                middleEastSAP_S.setPostingDate(middleEC.getDocumentDate());
                                middleEastSAP_S.setPeriod(middleEC.getDocumentDate().substring(0, i));
                                middleEastSAP_S.setCurrencyRate(middleEC.getCurrency());
                                String mm = middleEC.getDocumentDate().substring(0, i);
                                if ("01".equals(mm)) {
                                    mm = "12";
                                } else {
                                    mm = String.valueOf(Integer.parseInt(mm) - 1);
                                }
                                SimpleDateFormat sdf = new SimpleDateFormat("MM");
                                DateFormat df = new SimpleDateFormat("MMM", Locale.ENGLISH);
                                mm = df.format(sdf.parse(mm));
                                middleEastSAP_S.setReference("PAY SALARY" + " " + mm);
                                middleEastSAP_S.setDocHeaderText("PAY SALARY" + " " + mm);
                                middleEastSAP_S.setPostKey("50");
                                middleEastSAP_S.setGl(middleEC.getBank().substring(middleEC.getBank().indexOf(" ") + 1));
                                middleEastSAP_S.setAmount(String.valueOf(charge1));
                                middleEastSAP_S.setValueDate(middleEC.getDocumentDate());
                                middleEastSAP_S.setAssignment("PAY SALARY" + " " + mm);
                                middleEastSAP_S.setText("PAY SALARY" + " " + mm);
                                middleEastSAP_S.setLongText("");
                                middleEastSAP_S.setReasonCode("150");

                                middleEastSAP_S1.setId(middleEC.getId());
                                middleEastSAP_S1.setNo(uuids);
                                middleEastSAP_S1.setDocumentDate(middleEC.getDocumentDate());
                                middleEastSAP_S1.setType("SA");
                                middleEastSAP_S1.setCompanyCode("6770");
                                middleEastSAP_S1.setPostingDate(middleEC.getDocumentDate());
                                middleEastSAP_S1.setPeriod(middleEC.getDocumentDate().substring(0, i));
                                middleEastSAP_S1.setCurrencyRate(middleEC.getCurrency());
                                middleEastSAP_S1.setReference("PAY SALARY" + " " + mm);
                                middleEastSAP_S1.setDocHeaderText("PAY SALARY" + " " + mm);
                                middleEastSAP_S1.setPostKey("40");
                                middleEastSAP_S1.setGl("2151010000");
                                middleEastSAP_S1.setAmount(String.valueOf(charge1));
                                middleEastSAP_S1.setAssignment("PAY SALARY" + " " + mm);
                                middleEastSAP_S1.setText("PAY SALARY" + " " + mm);

                                middleEastLcSAPList.add(middleEastSAP_S);
                                middleEastLcSAPList.add(middleEastSAP_S1);

//                                }
//                                if (charge1 == 60000){
//                                    //费用
//                                    MiddleEastSAP middleEastSAP_C = new MiddleEastSAP();
//                                    MiddleEastSAP middleEastSAP_C1 = new MiddleEastSAP();
//                                    String uuidc = UUID.randomUUID().toString().replaceAll("-","");
//                                    middleEastSAP_C.setNo(uuidc);
//                                    middleEastSAP_C.setId(middleEC.getId());
//                                    middleEastSAP_C.setDocumentDate(middleEC.getDocumentDate());
//                                    middleEastSAP_C.setType("KA");
//                                    middleEastSAP_C.setCompanyCode("6770");
//                                    middleEastSAP_C.setPostingDate(middleEC.getDocumentDate());
//                                    middleEastSAP_C.setPeriod(middleEC.getDocumentDate().substring(0,i));
//                                    middleEastSAP_C.setCurrencyRate(middleEC.getCurrency());
//                                    String mm = middleEC.getDocumentDate().substring(0,i);
//                                    if ("01".equals(mm)){
//                                        mm = "12";
//                                    }else {
//                                        mm =String.valueOf(Integer.parseInt(mm) -1) ;
//                                    }
//                                    SimpleDateFormat sdf = new SimpleDateFormat("MM");
//                                    DateFormat df = new SimpleDateFormat("MMM",Locale.ENGLISH);
//                                    mm = df.format(sdf.parse(mm));
//                                    middleEastSAP_C.setReference("AAMER" + " " + mm);//月份需要使用英文
//                                    middleEastSAP_C.setDocHeaderText("AAMER" + " " + mm);
//                                    middleEastSAP_C.setGl("2181021300");
//                                    middleEastSAP_C.setPostKey("40");
//                                    middleEastSAP_C.setAmount("60000");
//                                    middleEastSAP_C.setProfitCenter("6770001419");
//                                    middleEastSAP_C.setAssignment("AAMER" + " " + mm);
//                                    middleEastSAP_C.setText("AAMER" + " " + mm);
//
//                                    middleEastSAP_C1.setNo(uuidc);
//                                    middleEastSAP_C1.setId(middleEC.getId());
//                                    middleEastSAP_C1.setDocumentDate(middleEC.getDocumentDate());
//                                    middleEastSAP_C1.setType("KA");
//                                    middleEastSAP_C1.setCompanyCode("6770");
//                                    middleEastSAP_C1.setPostingDate(middleEC.getDocumentDate());
//                                    middleEastSAP_C1.setPeriod(middleEC.getDocumentDate().substring(0,i));
//                                    middleEastSAP_C1.setCurrencyRate(middleEC.getCurrency());
//                                    middleEastSAP_C1.setReference("AAMER" + " " + mm);//月份需要使用英文
//                                    middleEastSAP_C1.setDocHeaderText("AAMER" + " " + mm);
//                                    middleEastSAP_C1.setPostKey("31");
//                                    middleEastSAP_C1.setGl("V999998864");
//                                    middleEastSAP_C1.setAmount("60000");
//                                    middleEastSAP_C1.setBlinDate(middleEC.getDocumentDate());
//                                    middleEastSAP_C1.setAssignment("AAMER" + " " + mm);
//                                    middleEastSAP_C1.setText("AAMER" + " " + mm);
//
//                                    middleEastLcSAPList.add(middleEastSAP_C);
//                                    middleEastLcSAPList.add(middleEastSAP_C1);
//
//                                    //付款60000凭证
//                                    MiddleEastSAP middleEastSAP_P = new MiddleEastSAP();
//                                    MiddleEastSAP middleEastSAP_P1 = new MiddleEastSAP();
//                                    String uuidp = UUID.randomUUID().toString().replaceAll("-","");
//                                    middleEastSAP_P.setId(middleEC.getId());
//                                    middleEastSAP_P.setNo(uuidp);
//                                    middleEastSAP_P.setDocumentDate(middleEC.getDocumentDate());
//                                    middleEastSAP_P.setType("KZ");
//                                    middleEastSAP_P.setCompanyCode("6770");
//                                    middleEastSAP_P.setPostingDate(middleEC.getDocumentDate());
//                                    middleEastSAP_P.setPeriod(middleEC.getDocumentDate().substring(0,i));
//                                    middleEastSAP_P.setCurrencyRate(middleEC.getCurrency());
//                                    String date1 = middleEC.getDocumentDate();
//                                    SimpleDateFormat sdf1 = new SimpleDateFormat("MM/dd/yyy");
//                                    SimpleDateFormat sd1 = new SimpleDateFormat("MM/dd");
//                                    date1 = sd1.format(sdf1.parse(date1));
//                                    middleEastSAP_P.setReference("PAY AAMER" + " " + date1);
//                                    middleEastSAP_P.setDocHeaderText("PAY AAMER" + " " + date1);
//                                    middleEastSAP_P.setPostKey("50");
//                                    middleEastSAP_P.setGl(middleEC.getBank().substring(middleEC.getBank().indexOf(" ") +1 ));
//                                    middleEastSAP_P.setAmount("60000");
//                                    middleEastSAP_P.setValueDate(middleEC.getDocumentDate());
//                                    middleEastSAP_P.setAssignment("PAY AAMER" + " " + date1);
//                                    middleEastSAP_P.setText("PAY AAMER" + " " + date1);
//                                    middleEastSAP_P.setReasonCode("140");
//                                    middleEastLcSAPList.add(middleEastSAP_P);
//
//                                    middleEastSAP_P1.setId(middleEC.getId());
//                                    middleEastSAP_P1.setNo(uuidp);
//                                    middleEastSAP_P1.setDocumentDate(middleEC.getDocumentDate());
//                                    middleEastSAP_P1.setType("KZ");
//                                    middleEastSAP_P1.setCompanyCode("6770");
//                                    middleEastSAP_P1.setPostingDate(middleEC.getDocumentDate());
//                                    middleEastSAP_P1.setPeriod(middleEC.getDocumentDate().substring(0,i));
//                                    middleEastSAP_P1.setCurrencyRate(middleEC.getCurrency());
//                                    middleEastSAP_P1.setReference("PAY AAMER" + " "+ date1);
//                                    middleEastSAP_P1.setDocHeaderText("PAY AAMER" + " " + date1);
//                                    middleEastSAP_P1.setPostKey("21");
//                                    middleEastSAP_P1.setGl("V999998864");
//                                    middleEastSAP_P1.setAmount("60000");
//                                    middleEastSAP_P1.setBlinDate(middleEC.getDocumentDate());
//                                    middleEastSAP_P1.setAssignment("PAY AAMER" + " " + date1);
//                                    middleEastSAP_P1.setText("PAY AAMER" + " " + date1);
//                                    middleEastLcSAPList.add(middleEastSAP_P1);
//                                }
//                                if (charge1 > 60000){
//                                    //费用
//                                    MiddleEastSAP middleEastSAP_C = new MiddleEastSAP();
//                                    MiddleEastSAP middleEastSAP_C1 = new MiddleEastSAP();
//                                    String uuidc = UUID.randomUUID().toString().replaceAll("-","");
//                                    middleEastSAP_C.setNo(uuidc);
//                                    middleEastSAP_C.setId(middleEC.getId());
//                                    middleEastSAP_C.setDocumentDate(middleEC.getDocumentDate());
//                                    middleEastSAP_C.setType("KA");
//                                    middleEastSAP_C.setCompanyCode("6770");
//                                    middleEastSAP_C.setPostingDate(middleEC.getDocumentDate());
//                                    middleEastSAP_C.setPeriod(middleEC.getDocumentDate().substring(0,i));
//                                    middleEastSAP_C.setCurrencyRate(middleEC.getCurrency());
//                                    String mm = middleEC.getDocumentDate().substring(0,i);
//                                    if ("01".equals(mm)){
//                                        mm = "12";
//                                    }else {
//                                        mm =String.valueOf(Integer.parseInt(mm) -1) ;
//                                    }
//                                    SimpleDateFormat sdf = new SimpleDateFormat("MM");
//                                    DateFormat df = new SimpleDateFormat("MMM",Locale.ENGLISH);
//                                    mm = df.format(sdf.parse(mm));
//                                    middleEastSAP_C.setReference("AAMER" + " " + mm);//月份需要使用英文
//                                    middleEastSAP_C.setDocHeaderText("AAMER" + " " + mm);
//                                    middleEastSAP_C.setGl("2181021300");
//                                    middleEastSAP_C.setPostKey("40");
//                                    middleEastSAP_C.setAmount("60000");
//                                    middleEastSAP_C.setProfitCenter("6770001419");
//                                    middleEastSAP_C.setAssignment("AAMER" + " " + mm);
//                                    middleEastSAP_C.setText("AAMER" + " " + mm);
//
//                                    middleEastSAP_C1.setNo(uuidc);
//                                    middleEastSAP_C1.setId(middleEC.getId());
//                                    middleEastSAP_C1.setDocumentDate(middleEC.getDocumentDate());
//                                    middleEastSAP_C1.setType("KA");
//                                    middleEastSAP_C1.setCompanyCode("6770");
//                                    middleEastSAP_C1.setPostingDate(middleEC.getDocumentDate());
//                                    middleEastSAP_C1.setPeriod(middleEC.getDocumentDate().substring(0,i));
//                                    middleEastSAP_C1.setCurrencyRate(middleEC.getCurrency());
//                                    middleEastSAP_C1.setReference("AAMER" + " " + mm);//月份需要使用英文
//                                    middleEastSAP_C1.setDocHeaderText("AAMER" + " " + mm);
//                                    middleEastSAP_C1.setPostKey("31");
//                                    middleEastSAP_C1.setGl("V999998864");
//                                    middleEastSAP_C1.setAmount("60000");
//                                    middleEastSAP_C1.setBlinDate(middleEC.getDocumentDate());
//                                    middleEastSAP_C1.setAssignment("AAMER" + " " + mm);
//                                    middleEastSAP_C1.setText("AAMER" + " " + mm);
//
//                                    middleEastLcSAPList.add(middleEastSAP_C);
//                                    middleEastLcSAPList.add(middleEastSAP_C1);
//
//                                    //付款60000
//                                    MiddleEastSAP middleEastSAP_P = new MiddleEastSAP();
//                                    MiddleEastSAP middleEastSAP_P1 = new MiddleEastSAP();
//                                    String uuidp = UUID.randomUUID().toString().replaceAll("-","");
//                                    middleEastSAP_P.setId(middleEC.getId());
//                                    middleEastSAP_P.setNo(uuidp);
//                                    middleEastSAP_P.setDocumentDate(middleEC.getDocumentDate());
//                                    middleEastSAP_P.setType("KZ");
//                                    middleEastSAP_P.setCompanyCode("6770");
//                                    middleEastSAP_P.setPostingDate(middleEC.getDocumentDate());
//                                    middleEastSAP_P.setPeriod(middleEC.getDocumentDate().substring(0,i));
//                                    middleEastSAP_P.setCurrencyRate(middleEC.getCurrency());
//                                    String date1 = middleEC.getDocumentDate();
//                                    SimpleDateFormat sdf1 = new SimpleDateFormat("MM/dd/yyy");
//                                    SimpleDateFormat sd1 = new SimpleDateFormat("MM/dd");
//                                    date1 = sd1.format(sdf1.parse(date1));
//                                    middleEastSAP_P.setReference("PAY AAMER" + " " + date1);
//                                    middleEastSAP_P.setDocHeaderText("PAY AAMER" + " " + date1);
//                                    middleEastSAP_P.setPostKey("50");
//                                    middleEastSAP_P.setGl(middleEC.getBank().substring(middleEC.getBank().indexOf(" ") +1 ));
//                                    middleEastSAP_P.setAmount("60000");
//                                    middleEastSAP_P.setValueDate(middleEC.getDocumentDate());
//                                    middleEastSAP_P.setAssignment("PAY AAMER" + " " + date1);
//                                    middleEastSAP_P.setText("PAY AAMER" + " " + date1);
//                                    middleEastSAP_P.setReasonCode("140");
//                                    middleEastLcSAPList.add(middleEastSAP_P);
//
//                                    middleEastSAP_P1.setId(middleEC.getId());
//                                    middleEastSAP_P1.setNo(uuidp);
//                                    middleEastSAP_P1.setDocumentDate(middleEC.getDocumentDate());
//                                    middleEastSAP_P1.setType("KZ");
//                                    middleEastSAP_P1.setCompanyCode("6770");
//                                    middleEastSAP_P1.setPostingDate(middleEC.getDocumentDate());
//                                    middleEastSAP_P1.setPeriod(middleEC.getDocumentDate().substring(0,i));
//                                    middleEastSAP_P1.setCurrencyRate(middleEC.getCurrency());
//                                    middleEastSAP_P1.setReference("PAY AAMER" + " "+ date1);
//                                    middleEastSAP_P1.setDocHeaderText("PAY AAMER" + " " + date1);
//                                    middleEastSAP_P1.setPostKey("21");
//                                    middleEastSAP_P1.setGl("V999998864");
//                                    middleEastSAP_P1.setAmount("60000");
//                                    middleEastSAP_P1.setBlinDate(middleEC.getDocumentDate());
//                                    middleEastSAP_P1.setAssignment("PAY AAMER" + " " + date1);
//                                    middleEastSAP_P1.setText("PAY AAMER" + " " + date1);
//                                    middleEastLcSAPList.add(middleEastSAP_P1);
//
//                                    //付款（总额-60000）
//                                    MiddleEastSAP middleEastSAP_S = new MiddleEastSAP();
//                                    MiddleEastSAP middleEastSAP_S1 = new MiddleEastSAP();
//                                    String uuids = UUID.randomUUID().toString().replaceAll("-","");
//                                    middleEastSAP_S.setNo(uuids);
//                                    middleEastSAP_S.setId(middleEC.getId());
//                                    middleEastSAP_S.setDocumentDate(middleEC.getDocumentDate());
//                                    middleEastSAP_S.setType("SA");
//                                    middleEastSAP_S.setCompanyCode("6770");
//                                    middleEastSAP_S.setPostingDate(middleEC.getDocumentDate());
//                                    middleEastSAP_S.setPeriod(middleEC.getDocumentDate().substring(0,i));
//                                    middleEastSAP_S.setCurrencyRate(middleEC.getCurrency());
//                                    middleEastSAP_S.setReference("PAY SALARY" + " " + mm);
//                                    middleEastSAP_S.setDocHeaderText("PAY SALARY" + " " + mm);
//                                    middleEastSAP_S.setPostKey("50");
//                                    middleEastSAP_S.setGl(middleEC.getBank().substring(middleEC.getBank().indexOf(" ") + 1));
//                                    middleEastSAP_S.setAmount(String.valueOf(charge1));
//                                    middleEastSAP_S.setValueDate(middleEC.getDocumentDate());
//                                    middleEastSAP_S.setAssignment("PAY SALARY" + " " + mm);
//                                    middleEastSAP_S.setText("PAY SALARY" + " " + mm);
//                                    middleEastSAP_S.setLongText("");
//                                    middleEastSAP_S.setReasonCode("150");
//
//                                    middleEastSAP_S1.setNo(uuids);
//                                    middleEastSAP_S1.setId(middleEC.getId());
//                                    middleEastSAP_S1.setDocumentDate(middleEC.getDocumentDate());
//                                    middleEastSAP_S1.setType("SA");
//                                    middleEastSAP_S1.setCompanyCode("6770");
//                                    middleEastSAP_S1.setPostingDate(middleEC.getDocumentDate());
//                                    middleEastSAP_S1.setPeriod(middleEC.getDocumentDate().substring(0,i));
//                                    middleEastSAP_S1.setCurrencyRate(middleEC.getCurrency());
//                                    middleEastSAP_S1.setReference("PAY SALARY" + " " + mm);
//                                    middleEastSAP_S1.setDocHeaderText("PAY SALARY" + " " + mm);
//                                    middleEastSAP_S1.setPostKey("40");
//                                    middleEastSAP_S1.setGl("2151010000");
//                                    middleEastSAP_S1.setAmount(String.valueOf(charge1));
//                                    middleEastSAP_S1.setAssignment("PAY SALARY" + " " + mm);
//                                    middleEastSAP_S1.setText("PAY SALARY" + " " + mm);
//
//                                    middleEastLcSAPList.add(middleEastSAP_S);
//                                    middleEastLcSAPList.add(middleEastSAP_S1);
//                                }
                            }
                            if ("A08".equals(customerCode)) {
                                middleEastSAP.setType("SA");
                                middleEastSAP.setReference("Deposit" + " " + date);
                                middleEastSAP.setDocHeaderText("Deposit" + " " + date);
                                middleEastSAP.setPostKey("50");
                                middleEastSAP.setGl(gl);
                                middleEastSAP.setAmount(convertData(middleEC.getCharge()));
                                middleEastSAP.setValueDate(middleEC.getDocumentDate());
                                middleEastSAP.setText("Deposit" + " " + date);
                                middleEastSAP.setAssignment("Deposit" + " " + date);
                                middleEastSAP.setLongText("");
                                middleEastSAP.setReasonCode("362");

                                middleEastSAP1.setType("SA");
                                middleEastSAP1.setReference("Deposit" + " " + date);
                                middleEastSAP1.setDocHeaderText("Deposit" + " " + date);
                                middleEastSAP1.setPostKey("40");
                                middleEastSAP1.setGl("1002999998");
                                middleEastSAP1.setAmount(convertData(middleEC.getCharge()));
                                middleEastSAP1.setValueDate(middleEC.getDocumentDate());
                                middleEastSAP1.setAssignment("Deposit" + " " + date);
                                middleEastSAP1.setText("Deposit" + " " + date);
                                middleEastSAP1.setLongText("");
                                middleEastSAP1.setReasonCode("362");

                                middleEastLcSAPList.add(middleEastSAP);
                                middleEastLcSAPList.add(middleEastSAP1);
                            }
                        } else {
                            status = status + "未找到相应业务类型!";
                        }
                    } else {
                        status = status + "客户代码为空!";
                    }
                }
                middleEC.setSapIncomeNo(status);
            } else {
                middleEC.setSapIncomeNo("数据认领异常");
            }
        }
        //合并LC类型数据生成的SAP数据
        middleEastTtSAPList.addAll(middleEastLcSAPList);
        //生成F-51表
        if (middleEastF51TtSAPList.size() > 0) {
            for (int i = 0; i < middleEastF51TtSAPList.size(); i++) {
                for (int j = 1; j < middleEastF51TtSAPList.size(); j++) {
                    if (i != j) {
                        if (middleEastF51TtSAPList.get(i).getId().equals(middleEastF51TtSAPList.get(j).getId())) {
                            middleEastF51TtSAPList.get(j).setNo(middleEastF51TtSAPList.get(i).getNo());
                        }
                    }
                }

            }
            for (MiddleEastSAP middleEastSAP : middleEastF51TtSAPList) {
                System.out.println(middleEastSAP.getAmount());
            }
            System.out.println("51集合大小:" + middleEastF51TtSAPList.size());
            excelOutput51_AR(middleEastF51TtSAPList, F51Path);
        }
        //获取银行附言为汇入汇款解付的数据的SAP
        List<MiddleEastSAP> MEOtherSAPList = new ArrayList<>();
        Set sets = new HashSet();//定义set来标识单元格合并的数据
        //汇入汇款解付数据，多认领情况，需要将多认领的数据提出单独处理
        List<MiddleEastAll> MEOtherListIncome = new ArrayList<>();
        for (MiddleEastAll middleEast : MEOtherList) {
            MiddleEastSAP middleEastSAP = new MiddleEastSAP();
            MiddleEastSAP middleEastSAP1 = new MiddleEastSAP();
            String charge = middleEast.getCharge();
            String income = middleEast.getIncome();
            String customerCode = middleEast.getCustomerCode();
            String lCInvoiceNumber = middleEast.getSummary();
            String status = "";
            //扣款金额为空，收款金额不为空
            if ((charge == null || "".equals(charge)) && (income != null && !"".equals(income))) {
                if (customerCode != null && !"".equals(customerCode)) {
                    if (customerCode.contains("6000") && customerCode.indexOf("6000", 0) == 0 && customerCode.length() == 10) {
                        if ("".equals(middleEast.getCustomerName())) {
                            status = status + "客户名称为空!";
                        }
                        if ("".equals(middleEast.getRecognizedAmount())) {
                            status = status + "认领金额为空!";
                        }
                        if ("".equals(middleEast.getPi())) {
                            status = status + "PI为空!";
                        }
                        if ("".equals(middleEast.getProductCode())) {
                            status = status + "产品线为空!";
                        }
                        if ("".equals(middleEast.getPrepayment())) {
                            status = status + "预付比例为空!";
                        } else {
                            if (!middleEast.getPrepayment().contains("N") && !middleEast.getPrepayment().contains("Y")
                                    && !middleEast.getPrepayment().contains("尾款")) {
                                status = status + "预付比例不符合规则!";
                            }
                        }
                    } else {
                        status = status + "不属于RPA处理的业务类型";
                    }
                    if ("".equals(status)) {
                        List<MiddleEastAll> lists = map.get(middleEast.getId());
                        DecimalFormat df = new DecimalFormat("#.00");
                        if (lists != null && lists.size() > 0) {
                            double sum = 0;
                            for (MiddleEastAll middleEastAll : lists) {
                                if (!sets.contains(middleEastAll.getCustomerCode())) {
                                    sets.add(middleEastAll.getCustomerCode());
                                }
                                if (middleEastAll.getRecognizedAmount() != null && !"".equals(middleEastAll.getRecognizedAmount())) {
                                    sum = sum + Double.parseDouble(middleEastAll.getRecognizedAmount());
                                }
                            }
                            if (sets.size() == 1) {
                                sum = Double.parseDouble(df.format(sum));
                                Double incomeNum = Double.parseDouble(middleEast.getIncome());
                                if (sum != incomeNum) {
                                    status = status + "认领金额与收款金额不一致";
                                } else {
                                    //单元格合并的数据在历史数据集合存放一条就可
                                    if (set.add(middleEast.getId())) {
                                        MEOtherListIncome.add(middleEast);
                                    }
                                }
                            } else {
                                status = status + "客户代码不一致";
                            }
                        } else {
                            String uuid = UUID.randomUUID().toString().replaceAll("-", "");
                            middleEastSAP.setNo(uuid);
                            middleEastSAP.setId(middleEast.getId());
                            middleEastSAP.setDocumentDate(middleEast.getDocumentDate());
                            middleEastSAP.setType("DZ");
                            middleEastSAP.setCompanyCode("6770");
                            middleEastSAP.setPostingDate(middleEast.getDocumentDate());
                            int i = middleEast.getDocumentDate().indexOf("/", 0);
                            int j = middleEast.getDocumentDate().lastIndexOf("/");
                            String date1 = middleEast.getDocumentDate();
                            SimpleDateFormat sd = new SimpleDateFormat("MM/dd/yyy");
                            SimpleDateFormat sd1 = new SimpleDateFormat("dd/MM");
                            date1 = sd1.format(sd.parse(date1));
                            middleEastSAP.setPeriod(middleEast.getDocumentDate().substring(0, i));
                            middleEastSAP.setCurrencyRate(middleEast.getCurrency());
                            middleEastSAP.setReference(middleEast.getPi());
                            middleEastSAP.setDocHeaderText(middleEast.getCustomerName() + " " + date1);
                            middleEastSAP.setPostKey("40");
                            String gl = middleEast.getBank().substring(middleEast.getBank().indexOf(" ") + 1);
                            middleEastSAP.setGl(gl);
                            middleEastSAP.setAmount(convertData(middleEast.getIncome()));
                            middleEastSAP.setValueDate(middleEast.getDocumentDate());
                            middleEastSAP.setAssignment(middleEast.getCustomerName());
                            middleEastSAP.setText(middleEast.getPi());
                            middleEastSAP.setReasonCode("112");
                            MEOtherSAPList.add(middleEastSAP);

                            if (middleEast.getPrepayment().contains("N") || middleEast.getPrepayment().contains("尾款")) {
                                middleEastSAP1.setPostKey("15");
                                middleEastSAP1.setBlinDate(middleEast.getDocumentDate());
                            }
                            if (middleEast.getPrepayment().contains("Y")) {
                                middleEastSAP1.setPostKey("19");
                                middleEastSAP1.setSglInd(">");
                                middleEastSAP1.setTaxCode("X0");
                                middleEastSAP1.setDuOn(middleEast.getDocumentDate());
                            }

                            middleEastSAP1.setNo(uuid);
                            middleEastSAP1.setId(middleEast.getId());
                            middleEastSAP1.setDocumentDate(middleEast.getDocumentDate());
                            middleEastSAP1.setType("DZ");
                            middleEastSAP1.setCompanyCode("6770");
                            middleEastSAP1.setPostingDate(middleEast.getDocumentDate());
                            middleEastSAP1.setPeriod(middleEast.getDocumentDate().substring(0, i));
                            middleEastSAP1.setCurrencyRate(middleEast.getCurrency());
                            middleEastSAP1.setReference(middleEast.getPi());
                            middleEastSAP1.setDocHeaderText(middleEast.getCustomerName() + " " + date1);
                            middleEastSAP1.setAmount(convertData(middleEast.getIncome()));
                            middleEastSAP1.setGl(middleEast.getCustomerCode());
                            middleEastSAP1.setAssignment(middleEast.getCustomerName());
                            middleEastSAP1.setText(middleEast.getPi());

                            MEOtherSAPList.add(middleEastSAP1);
                        }
                    }
                } else {
                    status = status + "客户代码为空!";
                }
                middleEast.setSapIncomeNo(status);
            }
            //收款金额为空，扣款金额不为空
            if ((charge != null && !"".equals(charge)) && (income == null || "".equals(income))) {
                if (customerCode != null && !"".equals(customerCode)) {
                    if (customerCode.contains("6000") || customerCode.contains("A01")) {
                        if (lCInvoiceNumber.contains("123") && lCInvoiceNumber.indexOf("123", 0) == 0 && lCInvoiceNumber.contains(" ")) {
                            String[] vl = lCInvoiceNumber.split(" ");
                            List<String> list = Arrays.asList(vl);
                            middleEastSAP.setAssignment(list.get(list.size() - 1));
                            middleEastSAP.setText(list.get(list.size() - 1));
                            middleEastSAP.setLongText(list.get(list.size() - 1));

                            middleEastSAP1.setText(list.get(list.size() - 1));
                            middleEastSAP1.setAssignment(list.get(list.size() - 1));
                            middleEastSAP1.setLongText(list.get(list.size() - 1));
                        } else {
                            middleEastSAP.setAssignment(lCInvoiceNumber);
                            middleEastSAP.setText(lCInvoiceNumber);
                            middleEastSAP.setLongText(lCInvoiceNumber);

                            middleEastSAP1.setAssignment(lCInvoiceNumber);
                            middleEastSAP1.setLongText(lCInvoiceNumber);
                            middleEastSAP1.setText(lCInvoiceNumber);
                        }
                        String uuid = UUID.randomUUID().toString().replaceAll("-", "");
                        middleEastSAP.setId(middleEast.getId());
                        middleEastSAP.setNo(uuid);
                        middleEastSAP.setDocumentDate(middleEast.getDocumentDate());
                        middleEastSAP.setCompanyCode("6770");
                        middleEastSAP.setPostingDate(middleEast.getDocumentDate());
                        int i = middleEast.getDocumentDate().indexOf("/", 0);
                        String date = middleEast.getDocumentDate();
                        SimpleDateFormat sd = new SimpleDateFormat("MM/dd/yyy");
                        SimpleDateFormat sdfd = new SimpleDateFormat("dd/MM");
                        date = sdfd.format(sd.parse(date));
                        middleEastSAP.setPeriod(middleEast.getDocumentDate().substring(0, i));
                        middleEastSAP.setCurrencyRate(middleEast.getCurrency());
                        String gl = middleEast.getBank().substring(middleEast.getBank().indexOf(" ") + 1);

                        middleEastSAP1.setId(middleEast.getId());
                        middleEastSAP1.setNo(uuid);
                        middleEastSAP1.setDocumentDate(middleEast.getDocumentDate());
                        middleEastSAP1.setCompanyCode("6770");
                        middleEastSAP.setPostingDate(middleEast.getDocumentDate());
                        middleEastSAP1.setPeriod(middleEast.getDocumentDate().substring(0, i));
                        middleEastSAP1.setCurrencyRate(middleEast.getCurrency());
                        if (customerCode.contains("6000") && customerCode.indexOf("6000", 0) == 0 && customerCode.length() == 10) {
                            if (middleEast.getCustomerName() != null && !"".equals(middleEast.getCustomerName())) {
                                middleEastSAP.setType("SA");
                                middleEastSAP.setDocHeaderText("BANK CHARGE" + " " + date);
                                middleEastSAP.setPostKey("50");
                                middleEastSAP.setReference(middleEast.getCustomerCode());
                                middleEastSAP.setGl(gl);
                                middleEastSAP.setAmount(convertData(middleEast.getCharge()));
                                middleEastSAP.setValueDate(middleEast.getDocumentDate());
                                middleEastSAP.setReasonCode("179");

                                middleEastSAP1.setType("SA");
                                middleEastSAP1.setPostKey("40");
                                middleEastSAP1.setDocHeaderText("BANK CHARGE" + " " + date);
                                middleEastSAP1.setReference(middleEast.getCustomerCode());
                                middleEastSAP1.setCostCenter(getCostCenter(middleEast.getProductCode(), filePath_Cost));
                                middleEastSAP1.setGl("5503040000");
                                middleEastSAP1.setAmount(convertData(middleEast.getCharge()));

                                if ("".equals(middleEastSAP1.getCostCenter()) && !"".equals(middleEast.getProductCode())) {
                                    status = status + "ProductCode填写有误,CostCenter无法正常匹配!";
                                } else {
                                    MEOtherSAPList.add(middleEastSAP);
                                    MEOtherSAPList.add(middleEastSAP1);
                                }

                            } else {
                                status = status + "客户名称为空!";
                            }
                        }
                        //客户代码为“Common”或者“手续费”
                        if ("A01".equals(customerCode)) {
                            middleEastSAP.setType("SA");
                            middleEastSAP.setPostKey("50");
                            middleEastSAP.setReference("BANK CHARGE");
                            middleEastSAP.setDocHeaderText("BANK CHARGE" + " " + date);
                            middleEastSAP.setGl(gl);
                            middleEastSAP.setAmount(convertData(middleEast.getCharge()));
                            middleEastSAP.setValueDate(middleEast.getDocumentDate());
                            middleEastSAP.setReasonCode("179");

                            middleEastSAP1.setType("SA");
                            middleEastSAP1.setReference("BANK CHARGE");
                            middleEastSAP1.setDocHeaderText("BANK CHARGE" + " " + date);
                            middleEastSAP1.setPostKey("40");
                            middleEastSAP1.setCostCenter("6677110102");
                            middleEastSAP1.setGl("5503040000");
                            middleEastSAP1.setAmount(convertData(middleEast.getCharge()));
                            MEOtherSAPList.add(middleEastSAP);
                            MEOtherSAPList.add(middleEastSAP1);
                        }
                    } else {
                        status = status + "不属于RPA处理的业务类型";
                    }
                } else {
                    status = status + "客户代码为空!";
                }
                middleEast.setSapIncomeNo(status);
            }
        }
        //将处理汇入汇款解付数据的集合合并到处理TT类型数据的集合
        middleEastTtSAPList.addAll(MEOtherSAPList);

        //汇入汇款解付类型，多认领数据获取sap数据
        List<MiddleEastSAP> MEOtherListIncomeSAP = getMEOtherIncomeSap(MEOtherList, MEOtherListIncome, map);
        middleEastTtSAPList.addAll(MEOtherListIncomeSAP);

        for (MiddleEastAll middleEast : MiddleEastListNo) {//未认领数据状态修改
            for (MiddleEastAll middleEastLc : MiddleEastLcList) {
                if (middleEast.getId().equals(middleEastLc.getId())) {
                    middleEast.setSapIncomeNo(middleEastLc.getSapIncomeNo());
                }
            }
            for (MiddleEastAll middleEastOther : MEOtherList) {
                if (middleEast.getId().equals(middleEastOther.getId())) {
                    middleEast.setSapIncomeNo(middleEastOther.getSapIncomeNo());
                }
            }
        }
        //全量台账数据状态修改
        for (MiddleEastAll middleEast : MELedgerList) {
            for (MiddleEastAll middleEastNo : MiddleEastListTT) {
                if (middleEast.getId().equals(middleEastNo.getId())) {
                    middleEast.setSapIncomeNo(middleEastNo.getSapIncomeNo());
                }
            }
            for (MiddleEastAll middleEastNo : MiddleEastListNo) {
                if (middleEast.getId().equals(middleEastNo.getId())) {
                    middleEast.setSapIncomeNo(middleEastNo.getSapIncomeNo());
                }
            }
            for (MiddleEastAll middleEastOther : MEOtherList) {
                if (middleEast.getId().equals(middleEastOther.getId())) {
                    middleEast.setSapIncomeNo(middleEastOther.getSapIncomeNo());
                }
            }
        }
        //月末将未认领的数据以一次性客户数据做处理（该处要做月末的触发判断）
        if ("other".equals(result)) {
            List<MiddleEastSAP> MEOnceSAPList = new ArrayList<>();
//            System.out.println("输出台账条数：" + MELedgerList.size());
            for (MiddleEastAll middleEast : MELedgerList) {
                MiddleEastSAP middleEastSAP = new MiddleEastSAP();
                MiddleEastSAP middleEastSAP1 = new MiddleEastSAP();
                Pattern pp = Pattern.compile("[^0-9]");//判断incomeNo中是否包含10为数字编码
                String incomeNo = pp.matcher(middleEast.getSapIncomeNo()).replaceAll("").trim();
                if ("".equals(middleEast.getSapNoForOther()) && incomeNo.length() < 10) {
                    //收款不为空，扣款为空
                    if (!"".equals(middleEast.getIncome()) && "".equals(middleEast.getCharge())) {
                        String uuid = UUID.randomUUID().toString().replaceAll("-", "");
                        middleEastSAP.setNo(uuid);
                        middleEastSAP.setId(middleEast.getId());
                        middleEastSAP.setDocumentDate(middleEast.getDocumentDate());
                        middleEastSAP.setType("DZ");
                        middleEastSAP.setCompanyCode("6770");
                        middleEastSAP.setPostingDate(middleEast.getDocumentDate());
                        int i = middleEast.getDocumentDate().indexOf("/", 0);
                        middleEastSAP.setPeriod(middleEast.getDocumentDate().substring(0, i));
                        middleEastSAP.setCurrencyRate(middleEast.getCurrency());
                        middleEastSAP.setReference(middleEast.getSummary());
                        middleEastSAP.setDocHeaderText(middleEast.getSummary());
                        middleEastSAP.setPostKey("40");
                        String gl = middleEast.getBank().substring(middleEast.getBank().indexOf(" ") + 1);
                        middleEastSAP.setGl(gl);
                        middleEastSAP.setAmount(convertData(middleEast.getIncome()));
                        middleEastSAP.setValueDate(middleEast.getDocumentDate());
                        middleEastSAP.setAssignment(middleEast.getSummary());
                        middleEastSAP.setText(middleEast.getSummary());
                        middleEastSAP.setLongText(middleEastSAP.getText());
                        middleEastSAP.setReasonCode("112");
                        MEOnceSAPList.add(middleEastSAP);

                        middleEastSAP1.setNo(uuid);
                        middleEastSAP1.setId(middleEast.getId());
                        middleEastSAP1.setDocumentDate(middleEast.getDocumentDate());
                        middleEastSAP1.setType("DZ");
                        middleEastSAP1.setCompanyCode("6770");
                        middleEastSAP1.setPostingDate(middleEast.getDocumentDate());
                        middleEastSAP1.setPeriod(middleEast.getDocumentDate().substring(0, i));
                        middleEastSAP1.setCurrencyRate(middleEast.getCurrency());
                        middleEastSAP1.setReference(middleEast.getCustomerCode());
                        middleEastSAP1.setDocHeaderText(middleEast.getSummary());
                        middleEastSAP1.setPostKey("19");
                        middleEastSAP1.setGl("6000007049");
                        middleEastSAP1.setSglInd(">");
                        middleEastSAP1.setAmount(convertData(middleEast.getIncome()));
                        middleEastSAP1.setTaxCode("X0");
                        middleEastSAP1.setDuOn(middleEast.getDocumentDate());
                        middleEastSAP1.setAssignment(middleEast.getPi());
                        middleEastSAP1.setText(middleEast.getPi());
                        middleEastSAP1.setLongText(middleEastSAP1.getText());
                        MEOnceSAPList.add(middleEastSAP1);
                    }
                    //收款为空，扣款不为空
                    if ("".equals(middleEast.getIncome()) && !"".equals(middleEast.getCharge())) {
                        String uuid = UUID.randomUUID().toString().replaceAll("-", "");
                        middleEastSAP.setNo(uuid);
                        middleEastSAP.setId(middleEast.getId());
                        middleEastSAP.setDocumentDate(middleEast.getDocumentDate());
                        middleEastSAP.setType("SA");
                        middleEastSAP.setCompanyCode("6770");
                        middleEastSAP.setPostingDate(middleEast.getDocumentDate());
                        int i = middleEast.getDocumentDate().indexOf("/", 0);
                        middleEastSAP.setPeriod(middleEast.getDocumentDate().substring(0, i));
                        middleEastSAP.setCurrencyRate(middleEast.getCurrency());
                        middleEastSAP.setReference(middleEast.getSummary());
                        middleEastSAP.setDocHeaderText(middleEast.getSummary());
                        middleEastSAP.setPostKey("50");
                        String gl = middleEast.getBank().substring(middleEast.getBank().indexOf(" ") + 1);
                        middleEastSAP.setGl(gl);
                        middleEastSAP.setAmount(convertData(middleEast.getCharge()));
                        middleEastSAP.setValueDate(middleEast.getDocumentDate());
                        middleEastSAP.setAssignment(middleEast.getSummary());
                        middleEastSAP.setText(middleEast.getSummary());
                        middleEastSAP.setLongText(middleEastSAP.getText());
                        middleEastSAP.setReasonCode("179");
                        MEOnceSAPList.add(middleEastSAP);

                        middleEastSAP1.setNo(uuid);
                        middleEastSAP1.setId(middleEast.getId());
                        middleEastSAP1.setDocumentDate(middleEast.getDocumentDate());
                        middleEastSAP1.setType("SA");
                        middleEastSAP1.setCompanyCode("6770");
                        middleEastSAP1.setPostingDate(middleEast.getDocumentDate());
                        middleEastSAP1.setPeriod(middleEast.getDocumentDate().substring(0, i));
                        middleEastSAP1.setCurrencyRate(middleEast.getCurrency());
                        middleEastSAP1.setReference(middleEast.getCustomerCode());
                        middleEastSAP1.setDocHeaderText(middleEast.getSummary());
                        middleEastSAP1.setPostKey("40");
                        middleEastSAP1.setGl("5503040000");
                        middleEastSAP1.setAmount(convertData(middleEast.getCharge()));
                        middleEastSAP1.setCostCenter("6677110102");
                        middleEastSAP1.setAssignment(middleEast.getPi());
                        middleEastSAP1.setText(middleEast.getPi());
                        middleEastSAP1.setLongText(middleEastSAP1.getText());

                        MEOnceSAPList.add(middleEastSAP1);
                    }
                    middleEast.setSapIncomeNo("other");//标识为一次性客户类型数据
                }
            }
            //合并一次性客户数据生成的SAP数据
            middleEastTtSAPList.addAll(MEOnceSAPList);
        }
        //生成全量数据台账表
        excelOutput_Log(MELedgerList, filePath);

        //

        return middleEastTtSAPList;
    }

    /**
     * 汇入汇款解付类型数据生成SAP（FB01）数据
     *
     * @param MEOtherList
     * @param MEOtherListIncome
     * @param map
     */
    public static List getMEOtherIncomeSap(List<MiddleEastAll> MEOtherList, List<MiddleEastAll> MEOtherListIncome, Map<String, List<MiddleEastAll>> map) throws Exception {
        List<MiddleEastSAP> middleEastSAPList = new ArrayList<>();
        if (MEOtherListIncome.size() > 0) {
            for (MiddleEastAll middleEastAll : MEOtherListIncome) {
                List<MiddleEastAll> list = map.get(middleEastAll.getId());
                int index = 0;
                for (MiddleEastAll middleEastAll2 : MEOtherList) {
                    if (middleEastAll2.getId().equals(middleEastAll.getId()) && "".equals(middleEastAll2.getSapIncomeNo())) {
                        index = index + 1;
                    }
                }
                if (list.size() == index) {
                    MiddleEastSAP middleEastSAP = new MiddleEastSAP();
                    MiddleEastAll middleEast = list.get(0);
                    String uuid = UUID.randomUUID().toString().replaceAll("-", "");
                    middleEastSAP.setId(middleEast.getId());
                    middleEastSAP.setNo(uuid);
                    middleEastSAP.setDocumentDate(middleEast.getDocumentDate());
                    middleEastSAP.setType("DZ");
                    middleEastSAP.setCompanyCode("6770");
                    middleEastSAP.setPostingDate(middleEast.getDocumentDate());
                    int i = middleEast.getDocumentDate().indexOf("/", 0);
                    int j = middleEast.getDocumentDate().lastIndexOf("/");
                    String date1 = middleEast.getDocumentDate();
                    SimpleDateFormat sd = new SimpleDateFormat("MM/dd/yyy");
                    SimpleDateFormat sd1 = new SimpleDateFormat("dd/MM");
                    date1 = sd1.format(sd.parse(date1));
                    middleEastSAP.setPeriod(middleEast.getDocumentDate().substring(0, i));
                    middleEastSAP.setCurrencyRate(middleEast.getCurrency());
                    middleEastSAP.setReference(middleEast.getPi());
                    middleEastSAP.setDocHeaderText(middleEast.getCustomerName() + " " + date1);
                    middleEastSAP.setPostKey("40");
                    String gl = middleEast.getBank().substring(middleEast.getBank().indexOf(" ") + 1);
                    middleEastSAP.setGl(gl);
                    middleEastSAP.setAmount(convertData(middleEast.getCharge()));
                    middleEastSAP.setValueDate(middleEast.getDocumentDate());
                    middleEastSAP.setAssignment(middleEast.getCustomerCode());
                    middleEastSAP.setText(middleEast.getPi());
                    middleEastSAP.setReasonCode("112");
                    middleEastSAPList.add(middleEastSAP);
                    for (MiddleEastAll middleEasts : list) {
                        MiddleEastSAP middleEastSAP1 = new MiddleEastSAP();
                        if (middleEasts.getPrepayment().contains("N") || middleEasts.getPrepayment().contains("尾款")) {
                            middleEastSAP1.setPostKey("15");
                            middleEastSAP1.setBlinDate(middleEast.getDocumentDate());
                        }
                        if (middleEasts.getPrepayment().contains("Y")) {
                            middleEastSAP1.setPostKey("19");
                            middleEastSAP1.setTaxCode("X0");
                            middleEastSAP1.setSglInd(">");
                            middleEastSAP1.setDuOn(middleEast.getDocumentDate());
                        }

                        middleEastSAP1.setId(middleEast.getId());
                        middleEastSAP1.setNo(uuid);
                        middleEastSAP1.setDocumentDate(middleEasts.getDocumentDate());
                        middleEastSAP1.setType("DZ");
                        middleEastSAP1.setCompanyCode("6770");
                        middleEastSAP1.setPostingDate(middleEasts.getDocumentDate());
                        middleEastSAP1.setPeriod(middleEasts.getDocumentDate().substring(0, i));
                        middleEastSAP1.setCurrencyRate(middleEasts.getCurrency());
                        middleEastSAP1.setReference(middleEasts.getPi());
                        middleEastSAP1.setDocHeaderText(middleEasts.getCustomerName() + " " + date1);
                        middleEastSAP1.setAmount(convertData(middleEasts.getCharge()));
                        middleEastSAP1.setGl(middleEasts.getCustomerCode());
                        middleEastSAP1.setAssignment(middleEasts.getCustomerCode());
                        middleEastSAP1.setText(middleEasts.getPi());

                        middleEastSAPList.add(middleEastSAP1);
                    }
                }
            }
        }
        return middleEastSAPList;
    }

    /**
     * TT类型数据生成SAP表
     *
     * @param MEIncomeTtList
     * @param MENowChargeTtList
     * @param MEOthChargeTtList
     * @param map
     */
    public static Map<String, List> cleanTtDataSap(List<MiddleEastAll> MENowIncomeTtList, List<MiddleEastAll> MEIncomeTtList, List<MiddleEastAll> MENowChargeTtList, List<MiddleEastAll> MEOthChargeTtList, Map<String, List<MiddleEastAll>> map, String filePath_Cost) throws Exception {
        Map<String, List> MapTotal = new HashMap<>();
        List<MiddleEastSAP> middleEastSAPList = new ArrayList<>();
        String status = "";
        if (MEIncomeTtList.size() > 0) {
            for (MiddleEastAll middleEast : MENowIncomeTtList) {
                MiddleEastSAP middleEastSAP = new MiddleEastSAP();
                MiddleEastSAP middleEastSAP1 = new MiddleEastSAP();
                String uuid = UUID.randomUUID().toString().replaceAll("-", "");
                middleEastSAP.setId(middleEast.getId());
                middleEastSAP.setNo(uuid);
                middleEastSAP.setDocumentDate(middleEast.getDocumentDate());
                middleEastSAP.setType("DZ");
                middleEastSAP.setCompanyCode("6770");
                middleEastSAP.setPostingDate(middleEast.getDocumentDate());
                int i = middleEast.getDocumentDate().indexOf("/", 0);
                String date1 = middleEast.getDocumentDate();
                SimpleDateFormat sd = new SimpleDateFormat("MM/dd/yyy");
                SimpleDateFormat sd1 = new SimpleDateFormat("dd/MM");
                date1 = sd1.format(sd.parse(date1));
                middleEastSAP.setPeriod(middleEast.getDocumentDate().substring(0, i));
                middleEastSAP.setCurrencyRate(middleEast.getCurrency());
                middleEastSAP.setReference(middleEast.getPi());
                middleEastSAP.setDocHeaderText(middleEast.getCustomerName() + " " + date1);
                middleEastSAP.setPostKey("40");
                String gl = middleEast.getBank().substring(middleEast.getBank().indexOf(" ") + 1);
                middleEastSAP.setGl(gl);
                middleEastSAP.setAmount(convertData(middleEast.getIncome()));
                middleEastSAP.setValueDate(middleEast.getDocumentDate());
                middleEastSAP.setAssignment(middleEast.getCustomerName());
                middleEastSAP.setText(middleEast.getPi());
                middleEastSAP.setReasonCode("112");

                if (middleEast.getPrepayment().contains("N") || middleEast.getPrepayment().contains("尾款")) {
                    middleEastSAP1.setPostKey("15");
                    middleEastSAP1.setBlinDate(middleEast.getDocumentDate());
                }
                if (middleEast.getPrepayment().contains("Y")) {
                    middleEastSAP1.setPostKey("19");
                    middleEastSAP1.setSglInd(">");
                    middleEastSAP1.setTaxCode("X0");
                    middleEastSAP1.setDuOn(middleEast.getDocumentDate());
                }
                middleEastSAP1.setId(middleEast.getId());
                middleEastSAP1.setNo(uuid);
                middleEastSAP1.setDocumentDate(middleEast.getDocumentDate());
                middleEastSAP1.setType("DZ");
                middleEastSAP1.setCompanyCode("6770");
                middleEastSAP1.setPostingDate(middleEast.getDocumentDate());
                middleEastSAP1.setPeriod(middleEast.getDocumentDate().substring(0, i));
                middleEastSAP1.setCurrencyRate(middleEast.getCurrency());
                middleEastSAP1.setReference(middleEast.getPi());
                middleEastSAP1.setDocHeaderText(middleEast.getCustomerName() + " " + date1);
                middleEastSAP1.setAmount(convertData(middleEast.getIncome()));
                middleEastSAP1.setGl(middleEast.getCustomerCode());
                middleEastSAP1.setAssignment(middleEast.getCustomerName());
                middleEastSAP1.setText(middleEast.getPi());

//                收款的时候也需要校验ProductCode能否匹配出

                String CostCenter = getCostCenter(middleEast.getProductCode(), filePath_Cost);
                if ("".equals(CostCenter) && !"".equals(middleEast.getProductCode())) {
                    status = status + "ProductCode填写有误,CostCenter无法正常匹配!";
                    middleEast.setSapIncomeNo(status);
                } else {
                    middleEastSAPList.add(middleEastSAP);
                    middleEastSAPList.add(middleEastSAP1);
                }
            }

        }
        if (MENowChargeTtList.size() > 0) {
            for (MiddleEastAll middleEast : MENowChargeTtList) {
                MiddleEastSAP middleEastSAP = new MiddleEastSAP();
                MiddleEastSAP middleEastSAP1 = new MiddleEastSAP();
                String uuid = UUID.randomUUID().toString().replaceAll("-", "");
                middleEastSAP.setId(middleEast.getId());
                middleEastSAP.setNo(uuid);
                middleEastSAP.setDocumentDate(middleEast.getDocumentDate());
                middleEastSAP.setType("SA");
                middleEastSAP.setCompanyCode("6770");
                middleEastSAP.setPostingDate(middleEast.getDocumentDate());
                int i = middleEast.getDocumentDate().indexOf("/", 0);
                int j = middleEast.getDocumentDate().lastIndexOf("/");
                String date1 = middleEast.getDocumentDate();
                SimpleDateFormat sd = new SimpleDateFormat("MM/dd/yyy");
                SimpleDateFormat sd1 = new SimpleDateFormat("dd/MM");
                date1 = sd1.format(sd.parse(date1));
                middleEastSAP.setPeriod(middleEast.getDocumentDate().substring(0, i));
                middleEastSAP.setCurrencyRate(middleEast.getCurrency());
                middleEastSAP.setReference(middleEast.getCustomerCode());
                middleEastSAP.setDocHeaderText("BANK CHARGE" + " " + date1);
                middleEastSAP.setPostKey("50");
                String gl = middleEast.getBank().substring(middleEast.getBank().indexOf(" ") + 1);
                middleEastSAP.setGl(gl);
                middleEastSAP.setAmount(convertData(middleEast.getCharge()));
                middleEastSAP.setValueDate(middleEast.getDocumentDate());
                middleEastSAP.setAssignment(middleEast.getPi());
                middleEastSAP.setText(middleEast.getPi());
                middleEastSAP.setReasonCode("179");

                middleEastSAP1.setId(middleEast.getId());
                middleEastSAP1.setNo(uuid);
                middleEastSAP1.setDocumentDate(middleEast.getDocumentDate());
                middleEastSAP1.setType("SA");
                middleEastSAP1.setCompanyCode("6770");
                middleEastSAP1.setPostingDate(middleEast.getDocumentDate());
                middleEastSAP1.setPeriod(middleEast.getDocumentDate().substring(0, i));
                middleEastSAP1.setCurrencyRate(middleEast.getCurrency());
                middleEastSAP1.setReference(middleEast.getCustomerCode());
                middleEastSAP1.setDocHeaderText("BANK CHARGE" + " " + date1);
                middleEastSAP1.setPostKey("40");
                middleEastSAP1.setGl("5503040000");
                middleEastSAP1.setAmount(convertData(middleEast.getCharge()));
                middleEastSAP1.setCostCenter(getCostCenter(middleEast.getProductCode(), filePath_Cost));
                middleEastSAP1.setAssignment(middleEast.getPi());
                middleEastSAP1.setText(middleEast.getPi());
                if ("".equals(middleEastSAP1.getCostCenter()) && !"".equals(middleEast.getProductCode())) {
                    status = status + "ProductCode填写有误,CostCenter无法正常匹配!";
                    middleEast.setSapIncomeNo(status);
                } else {
                    middleEastSAPList.add(middleEastSAP);
                    middleEastSAPList.add(middleEastSAP1);
                }
            }
        }
        if (MEOthChargeTtList.size() > 0) {
            for (MiddleEastAll middleEast : MEOthChargeTtList) {
                MiddleEastSAP middleEastSAP = new MiddleEastSAP();
                MiddleEastSAP middleEastSAP1 = new MiddleEastSAP();
                String uuid = UUID.randomUUID().toString().replaceAll("-", "");
                middleEastSAP.setId(middleEast.getId());
                middleEastSAP.setNo(uuid);
                middleEastSAP.setDocumentDate(middleEast.getDocumentDate());
                middleEastSAP.setType("SA");
                middleEastSAP.setCompanyCode("6770");
                middleEastSAP.setPostingDate(middleEast.getDocumentDate());
                int i = middleEast.getDocumentDate().indexOf("/", 0);
                int j = middleEast.getDocumentDate().lastIndexOf("/");
                String date1 = middleEast.getDocumentDate();
                SimpleDateFormat sd = new SimpleDateFormat("MM/dd/yyy");
                SimpleDateFormat sd1 = new SimpleDateFormat("dd/MM");
                date1 = sd1.format(sd.parse(date1));
                middleEastSAP.setPeriod(middleEast.getDocumentDate().substring(0, i));
                middleEastSAP.setCurrencyRate(middleEast.getCurrency());
                middleEastSAP.setReference("BANK CHARGE");
                middleEastSAP.setDocHeaderText("BANK CHARGE" + " " + date1);
                middleEastSAP.setPostKey("50");
                String gl = middleEast.getBank().substring(middleEast.getBank().indexOf(" ") + 1);
                middleEastSAP.setGl(gl);
                middleEastSAP.setAmount(convertData(middleEast.getCharge()));
                middleEastSAP.setValueDate(middleEast.getDocumentDate());
                middleEastSAP.setAssignment("BANK CHARGE");
                middleEastSAP.setText("BANK CHARGE");
                middleEastSAP.setReasonCode("179");
                middleEastSAPList.add(middleEastSAP);

                middleEastSAP1.setId(middleEast.getId());
                middleEastSAP1.setNo(uuid);
                middleEastSAP1.setDocumentDate(middleEast.getDocumentDate());
                middleEastSAP1.setType("SA");
                middleEastSAP1.setCompanyCode("6770");
                middleEastSAP1.setPostingDate(middleEast.getDocumentDate());
                middleEastSAP1.setPeriod(middleEast.getDocumentDate().substring(0, i));
                middleEastSAP1.setCurrencyRate(middleEast.getCurrency());
                middleEastSAP1.setReference("BANK CHARGE");
                middleEastSAP1.setDocHeaderText("BANK CHARGE" + " " + date1);
                middleEastSAP1.setPostKey("40");
                middleEastSAP1.setGl("5503040000");
                middleEastSAP1.setAmount(convertData(middleEast.getCharge()));
                middleEastSAP1.setCostCenter("6677110102");
                middleEastSAP1.setAssignment("BANK CHARGE");
                middleEastSAP1.setText("BANK CHARGE");
                middleEastSAPList.add(middleEastSAP1);
            }
        }
            //多人认领的情况---------------------------------------------------------------------------------------------------------
            //单元格合并的数据需要单独处理
            //遍历map 取出合并单元格数据的key,

            for (String key : map.keySet()) {
                int num = 0;
                //遍历所有的正数数据集合，统计ID与合并单元格数据ID相等的数据的数量
                for (MiddleEastAll middleEast : MEIncomeTtList) {
                    if (key.equals(middleEast.getId()) && ("".equals(middleEast.getSapIncomeNo()) || middleEast.getSapIncomeNo() == null)) {
                        num = num + 1;
                    }
                }
                List<MiddleEastAll> lists = map.get(key);
                if (num == lists.size()) {
                    MiddleEastSAP middleEastSAP = new MiddleEastSAP();
                    MiddleEastAll middleEast = lists.get(0);
                    String uuid = UUID.randomUUID().toString().replaceAll("-", "");
                    DateFormat dFormat = new SimpleDateFormat("MM/dd/yyyy");
                    Calendar calendar = Calendar.getInstance();
                    calendar.setTime(new Date());
                    int i = middleEast.getDocumentDate().indexOf("/", 0);
                    //日期 + 时间
                    SimpleDateFormat sd1 = new SimpleDateFormat("dd/MM");
                    String date = sd1.format(calendar.getTime());
                    calendar.set(Calendar.DAY_OF_MONTH, 1);//当前日期既为本月第一天
                    System.out.println(middleEast.getSapNoForOther());
                    if (middleEast.getSapNoForOther().length()==0){
                        middleEastSAP.setId(middleEast.getId());
                        middleEastSAP.setNo(uuid);
                        middleEastSAP.setDocumentDate(middleEast.getDocumentDate());
                        middleEastSAP.setType("DZ");
                        middleEastSAP.setCompanyCode("6770");
////                //本月第一天
                        middleEastSAP.setPostingDate(dFormat.format(calendar.getTime()));
////                //当前月份
                        String month = String.valueOf(calendar.get(Calendar.MONTH) + 1);
////
                        middleEastSAP.setId(middleEast.getId());
                        middleEastSAP.setNo(uuid);
                        middleEastSAP.setDocumentDate(middleEast.getDocumentDate());
                        middleEastSAP.setType("DZ");
                        middleEastSAP.setCompanyCode("6770");
                        middleEastSAP.setPostingDate(middleEast.getDocumentDate());
                        String date1 = middleEast.getDocumentDate();
                        SimpleDateFormat sd = new SimpleDateFormat("MM/dd/yyy");
                        date1 = sd1.format(sd.parse(date1));
                        middleEastSAP.setPeriod(middleEast.getDocumentDate().substring(0, i));
                        middleEastSAP.setCurrencyRate(middleEast.getCurrency());
                        middleEastSAP.setReference(middleEast.getPi());
                        middleEastSAP.setDocHeaderText(middleEast.getCustomerName() + " " + date1);
                        middleEastSAP.setPostKey("40");
                        String gl = middleEast.getBank().substring(middleEast.getBank().indexOf(" ") + 1);
                        middleEastSAP.setGl(gl);
                        middleEastSAP.setAmount(convertData(middleEast.getIncome()));
                        middleEastSAP.setValueDate(middleEast.getDocumentDate());
                        middleEastSAP.setAssignment(middleEast.getCustomerName());
                        middleEastSAP.setText(middleEast.getPi());
                        middleEastSAP.setReasonCode("112");
                        middleEastSAPList.add(middleEastSAP);
                    }else {
                        if (!(middleEast.getSapNoForOther().substring(0, 2).equals("14") && (middleEast.getSapClearingNo() == null || middleEast.getSapClearingNo().equals("")))) {
                            middleEastSAP.setId(middleEast.getId());
                            middleEastSAP.setNo(uuid);
                            middleEastSAP.setDocumentDate(middleEast.getDocumentDate());
                            middleEastSAP.setType("DZ");
                            middleEastSAP.setCompanyCode("6770");
////                //本月第一天
                            middleEastSAP.setPostingDate(dFormat.format(calendar.getTime()));
////                //当前月份
                            String month = String.valueOf(calendar.get(Calendar.MONTH) + 1);
////
                            middleEastSAP.setId(middleEast.getId());
                            middleEastSAP.setNo(uuid);
                            middleEastSAP.setDocumentDate(middleEast.getDocumentDate());
                            middleEastSAP.setType("DZ");
                            middleEastSAP.setCompanyCode("6770");
                            middleEastSAP.setPostingDate(middleEast.getDocumentDate());
                            String date1 = middleEast.getDocumentDate();
                            SimpleDateFormat sd = new SimpleDateFormat("MM/dd/yyy");
                            date1 = sd1.format(sd.parse(date1));
                            middleEastSAP.setPeriod(middleEast.getDocumentDate().substring(0, i));
                            middleEastSAP.setCurrencyRate(middleEast.getCurrency());
                            middleEastSAP.setReference(middleEast.getPi());
                            middleEastSAP.setDocHeaderText(middleEast.getCustomerName() + " " + date1);
                            middleEastSAP.setPostKey("40");
                            String gl = middleEast.getBank().substring(middleEast.getBank().indexOf(" ") + 1);
                            middleEastSAP.setGl(gl);
                            middleEastSAP.setAmount(convertData(middleEast.getIncome()));
                            middleEastSAP.setValueDate(middleEast.getDocumentDate());
                            middleEastSAP.setAssignment(middleEast.getCustomerName());
                            middleEastSAP.setText(middleEast.getPi());
                            middleEastSAP.setReasonCode("112");
                            middleEastSAPList.add(middleEastSAP);
                        }
                    }



                    for (MiddleEastAll middleEasts : lists) {
                        if (middleEasts.getSapNoForOther().length()==0){
                            MiddleEastSAP middleEastSAP1 = new MiddleEastSAP();
                            if (middleEasts.getPrepayment().contains("N") || middleEasts.getPrepayment().contains("尾款")) {
                                middleEastSAP1.setPostKey("15");
                                middleEastSAP1.setBlinDate(middleEasts.getDocumentDate());
                            }
                            if (middleEasts.getPrepayment().contains("Y")) {
                                middleEastSAP1.setPostKey("19");
                                middleEastSAP1.setTaxCode("X0");
                                middleEastSAP1.setSglInd(">");
                                middleEastSAP1.setDuOn(middleEasts.getDocumentDate());
                            }
                            middleEastSAP1.setId(middleEasts.getId());
                            middleEastSAP1.setNo(uuid);
                            middleEastSAP1.setDocumentDate(middleEasts.getDocumentDate());
                            middleEastSAP1.setType("DZ");
                            middleEastSAP1.setCompanyCode("6770");
                            middleEastSAP1.setPostingDate(dFormat.format(calendar.getTime()));
                            middleEastSAP1.setPeriod(middleEast.getDocumentDate().substring(0, i));
                            middleEastSAP1.setCurrencyRate(middleEasts.getCurrency());
                            middleEastSAP1.setReference(middleEasts.getPi());
                            middleEastSAP1.setDocHeaderText(middleEasts.getCustomerName() + " " + date);
                            middleEastSAP1.setAmount(convertData(middleEasts.getRecognizedAmount()));
                            middleEastSAP1.setGl(middleEasts.getCustomerCode());
                            middleEastSAP1.setAssignment(middleEasts.getCustomerName());
                            middleEastSAP1.setText(middleEasts.getPi());
                            middleEastSAPList.add(middleEastSAP1);
                        }else {
                            if (!(middleEasts.getSapNoForOther().substring(0, 2).equals("14") && (middleEasts.getSapClearingNo() == null || middleEasts.getSapClearingNo().equals("")))) {

                                MiddleEastSAP middleEastSAP1 = new MiddleEastSAP();
                                if (middleEasts.getPrepayment().contains("N") || middleEasts.getPrepayment().contains("尾款")) {
                                    middleEastSAP1.setPostKey("15");
                                    middleEastSAP1.setBlinDate(middleEasts.getDocumentDate());
                                }
                                if (middleEasts.getPrepayment().contains("Y")) {
                                    middleEastSAP1.setPostKey("19");
                                    middleEastSAP1.setTaxCode("X0");
                                    middleEastSAP1.setSglInd(">");
                                    middleEastSAP1.setDuOn(middleEasts.getDocumentDate());
                                }
                                middleEastSAP1.setId(middleEasts.getId());
                                middleEastSAP1.setNo(uuid);
                                middleEastSAP1.setDocumentDate(middleEasts.getDocumentDate());
                                middleEastSAP1.setType("DZ");
                                middleEastSAP1.setCompanyCode("6770");
                                middleEastSAP1.setPostingDate(dFormat.format(calendar.getTime()));
                                middleEastSAP1.setPeriod(middleEast.getDocumentDate().substring(0, i));
                                middleEastSAP1.setCurrencyRate(middleEasts.getCurrency());
                                middleEastSAP1.setReference(middleEasts.getPi());
                                middleEastSAP1.setDocHeaderText(middleEasts.getCustomerName() + " " + date);
                                middleEastSAP1.setAmount(convertData(middleEasts.getRecognizedAmount()));
                                middleEastSAP1.setGl(middleEasts.getCustomerCode());
                                middleEastSAP1.setAssignment(middleEasts.getCustomerName());
                                middleEastSAP1.setText(middleEasts.getPi());
                                middleEastSAPList.add(middleEastSAP1);
                            }
                        }

                    }
                }
            }

        MapTotal.put("List", middleEastSAPList);
        MapTotal.put("ledger", MENowIncomeTtList);
        return MapTotal;
    }

    /**
     * 一次性客户类型的，TT类型数据
     *
     * @param MENowIncomeTtList
     * @param MEIncomeTtList
     * @param MENowChargeTtList
     * @param map
     * @param filePath_Cost
     * @return
     * @throws Exception
     */
    public static Map<String, List> cleanForOtherTtDataSap(List<MiddleEastAll> MENowIncomeTtList, List<MiddleEastAll> MEIncomeTtList, List<MiddleEastAll> MENowChargeTtList, Map<String, List<MiddleEastAll>> map, String filePath_Cost) throws Exception {
        List<MiddleEastSAP> middleEastSAPList = new ArrayList<>();
        List<MiddleEastSAP> middleEastSAP_F_51_List = new ArrayList<>();
        Map<String, List> final_map = new HashMap<>();
        if (MEIncomeTtList.size() > 0) {
            for (MiddleEastAll middleEast : MENowIncomeTtList) {
                MiddleEastSAP middleEastSAP = new MiddleEastSAP();
                MiddleEastSAP middleEastSAP1 = new MiddleEastSAP();
                String uuid = UUID.randomUUID().toString().replaceAll("-", "");
                middleEastSAP.setId(middleEast.getId());
                middleEastSAP.setNo(uuid);
                middleEastSAP.setDocumentDate(middleEast.getDocumentDate());
                middleEastSAP.setType("DZ");
                middleEastSAP.setCompanyCode("6770");
                DateFormat dFormat = new SimpleDateFormat("MM/dd/yyyy");
                Calendar calendar = Calendar.getInstance();
                calendar.setTime(new Date());
                calendar.set(Calendar.DAY_OF_MONTH, 1);//当前日期既为本月第一天
                //本月第一天
                middleEastSAP.setPostingDate(dFormat.format(calendar.getTime()));
                //当前月份
                String month = String.valueOf(calendar.get(Calendar.MONTH) + 1);
                //日期 + 时间
                SimpleDateFormat sd1 = new SimpleDateFormat("dd/MM");
                String date = sd1.format(calendar.getTime());
                middleEastSAP.setPeriod(month);
                middleEastSAP.setCurrencyRate(middleEast.getCurrency());
                middleEastSAP.setReference(middleEast.getPi());
                middleEastSAP.setDocHeaderText(middleEast.getCustomerName() + " " + date);
                middleEastSAP.setPostKey("09");
                middleEastSAP.setGl("6000007049");
                middleEastSAP.setSglInd(">");
                middleEastSAP.setAmount(convertData(middleEast.getIncome()));
                middleEastSAP.setTaxCode("X0");
                middleEastSAP.setDuOn(middleEast.getDocumentDate());
                middleEastSAP.setAssignment(middleEast.getCustomerName());
                middleEastSAP.setText(middleEast.getPi());
                if (!(middleEast.getSapNoForOther().substring(0, 2).equals("14") && (middleEast.getSapClearingNo() == null || middleEast.getSapClearingNo().equals("")) && !("A01".equals(middleEast.getCustomerCode())) )) {
                    if (("A01".equals(middleEast.getCustomerCode())) &&  ((middleEast.getSapNoForOther()!=null || !("".equals(middleEast.getSapNoForOther()))))){

                    }else {
                        middleEastSAPList.add(middleEastSAP);
                    }
                }

                if (middleEast.getPrepayment().contains("N") || middleEast.getPrepayment().contains("尾款")) {
                    middleEastSAP1.setPostKey("15");
                    middleEastSAP1.setBlinDate(middleEast.getDocumentDate());
                }
                if (middleEast.getPrepayment().contains("Y")) {
                    middleEastSAP1.setPostKey("19");
                    middleEastSAP1.setSglInd(">");
                    middleEastSAP1.setTaxCode("");
                    middleEastSAP1.setDuOn(middleEast.getDocumentDate());
                }
                middleEastSAP1.setId(middleEast.getId());
                middleEastSAP1.setNo(uuid);
                middleEastSAP1.setDocumentDate(middleEast.getDocumentDate());
                middleEastSAP1.setType("DZ");
                middleEastSAP1.setCompanyCode("6770");
                middleEastSAP1.setPostingDate(dFormat.format(calendar.getTime()));
                middleEastSAP1.setPeriod(month);
                middleEastSAP1.setCurrencyRate(middleEast.getCurrency());
                middleEastSAP1.setReference(middleEast.getPi());
                middleEastSAP1.setDocHeaderText(middleEast.getCustomerName() + " " + date);
                middleEastSAP1.setAmount(convertData(middleEast.getIncome()));
                middleEastSAP1.setGl(middleEast.getCustomerCode());
                middleEastSAP1.setAssignment(middleEast.getCustomerName());
                middleEastSAP1.setText(middleEast.getPi());
                middleEastSAP1.setSapforother(middleEast.getSapNoForOther());
                if (!(middleEast.getSapNoForOther().substring(0, 2).equals("14") && (middleEast.getSapClearingNo() == null || middleEast.getSapClearingNo().equals("")) && !("A01".equals(middleEast.getCustomerCode())))) {
                    if (("A01".equals(middleEast.getCustomerCode())) &&  ((middleEast.getSapNoForOther()!=null || !("".equals(middleEast.getSapNoForOther())))) && (middleEast.getSapIncomeNo().equals(middleEast.getSapNoForOther()))){

                    }else {
                        middleEastSAPList.add(middleEastSAP);
                    }
                } else {
                    middleEastSAP_F_51_List.add(middleEastSAP1);
                }
            }
        }
        if (MENowChargeTtList.size() > 0) {
            for (MiddleEastAll middleEast : MENowChargeTtList) {
                MiddleEastSAP middleEastSAP = new MiddleEastSAP();
                MiddleEastSAP middleEastSAP1 = new MiddleEastSAP();
                String uuid = UUID.randomUUID().toString().replaceAll("-", "");
                middleEastSAP.setId(middleEast.getId());
                middleEastSAP.setNo(uuid);
                middleEastSAP.setDocumentDate(middleEast.getDocumentDate());
                middleEastSAP.setType("SA");
                middleEastSAP.setCompanyCode("6770");

                DateFormat dFormat = new SimpleDateFormat("MM/dd/yyyy");
                Calendar calendar = Calendar.getInstance();
                calendar.setTime(new Date());
                calendar.set(Calendar.DAY_OF_MONTH, 1);//当前日期既为本月第一天
                //本月第一天
                middleEastSAP.setPostingDate(dFormat.format(calendar.getTime()));
                //当前月份
                String month = String.valueOf(calendar.get(Calendar.MONTH) + 1);
                //日期 + 时间
                SimpleDateFormat sd1 = new SimpleDateFormat("dd/MM");
                String date = sd1.format(calendar.getTime());
                middleEastSAP.setPeriod(month);
                middleEastSAP.setCurrencyRate(middleEast.getCurrency());
                middleEastSAP.setReference(middleEast.getCustomerCode());
                middleEastSAP.setDocHeaderText("BANK CHARGE" + " " + date);
                middleEastSAP.setPostKey("50");
                middleEastSAP.setGl("5503040000");
                middleEastSAP.setAmount(convertData(middleEast.getCharge()));
                middleEastSAP.setCostCenter("6677110102");
                middleEastSAP.setAssignment(middleEast.getPi());
                middleEastSAP.setText(middleEast.getPi());


                middleEastSAP1.setId(middleEast.getId());
                middleEastSAP1.setNo(uuid);
                middleEastSAP1.setDocumentDate(middleEast.getDocumentDate());
                middleEastSAP1.setType("SA");
                middleEastSAP1.setCompanyCode("6770");
                middleEastSAP1.setPostingDate(dFormat.format(calendar.getTime()));
                middleEastSAP1.setPeriod(month);
                middleEastSAP1.setCurrencyRate(middleEast.getCurrency());
                middleEastSAP1.setReference(middleEast.getCustomerCode());
                middleEastSAP1.setDocHeaderText("BANK CHARGE" + " " + date);
                middleEastSAP1.setPostKey("40");
                middleEastSAP1.setGl("5503040000");
                middleEastSAP1.setAmount(convertData(middleEast.getCharge()));
                middleEastSAP1.setCostCenter(getCostCenter(middleEast.getProductCode(), filePath_Cost));
                middleEastSAP1.setAssignment(middleEast.getPi());
                middleEastSAP1.setText(middleEast.getPi());

                if ("".equals(middleEastSAP1.getCostCenter()) && !"".equals(middleEast.getProductCode())) {
//                    status = status + "ProductCode填写有误,CostCenter无法正常匹配!";
                } else {
                    middleEastSAPList.add(middleEastSAP);
                    middleEastSAPList.add(middleEastSAP1);
                }

            }
        }

        //单元格合并的数据需要单独处理
        //遍历map 取出合并单元格数据的key,
        for (String key : map.keySet()) {
            int num = 0;
            //遍历所有的正数数据集合，统计ID与合并单元格数据ID相等的数据的数量
            for (MiddleEastAll middleEast : MEIncomeTtList) {
                if (key.equals(middleEast.getId()) && !"".equals(middleEast.getSapNoForOther())) {
                    num = num + 1;
                }
            }
            List<MiddleEastAll> lists = map.get(key);
            if (num == lists.size()) {
                String uuid = UUID.randomUUID().toString().replaceAll("-", "");
                DateFormat dFormat = new SimpleDateFormat("MM/dd/yyyy");
                Calendar calendar = Calendar.getInstance();
                calendar.setTime(new Date());
                calendar.set(Calendar.DAY_OF_MONTH, 1);//当前日期既为本月第一天
//                //本月第一天
//                middleEastSAP.setPostingDate(dFormat.format(calendar.getTime()));
//                //当前月份
                String month = String.valueOf(calendar.get(Calendar.MONTH) + 1);
//                //日期 + 时间
                SimpleDateFormat sd1 = new SimpleDateFormat("dd/MM");
                String date = sd1.format(calendar.getTime());
                for (MiddleEastAll middleEasts : lists) {
                    MiddleEastSAP middleEastSAP1 = new MiddleEastSAP();
                    if (middleEasts.getPrepayment().contains("N") || middleEasts.getPrepayment().contains("尾款")) {
                        middleEastSAP1.setPostKey("15");
                        middleEastSAP1.setBlinDate(middleEasts.getDocumentDate());
                    }
                    if (middleEasts.getPrepayment().contains("Y")) {
                        middleEastSAP1.setPostKey("19");
                        middleEastSAP1.setTaxCode("X0");
                        middleEastSAP1.setSglInd(">");
                        middleEastSAP1.setDuOn(middleEasts.getDocumentDate());
                    }

                    middleEastSAP1.setId(middleEasts.getId());
                    middleEastSAP1.setNo(uuid);
                    middleEastSAP1.setDocumentDate(middleEasts.getDocumentDate());
                    middleEastSAP1.setType("DZ");
                    middleEastSAP1.setCompanyCode("6770");
                    middleEastSAP1.setPostingDate(dFormat.format(calendar.getTime()));
                    middleEastSAP1.setPeriod(month);
                    middleEastSAP1.setCurrencyRate(middleEasts.getCurrency());
                    middleEastSAP1.setReference(middleEasts.getPi());
                    middleEastSAP1.setDocHeaderText(middleEasts.getCustomerName() + " " + date);
                    middleEastSAP1.setAmount(convertData(middleEasts.getRecognizedAmount()));
                    middleEastSAP1.setGl(middleEasts.getCustomerCode());
                    middleEastSAP1.setAssignment(middleEasts.getCustomerName());
                    middleEastSAP1.setText(middleEasts.getPi());
                    middleEastSAP1.setSapforother(middleEasts.getSapNoForOther());

                    middleEastSAP_F_51_List.add(middleEastSAP1);
                }
            }
        }
        final_map.put("middleEastSAPList", middleEastSAPList);
        final_map.put("middleEastSAP_F_51_List", middleEastSAP_F_51_List);
        return final_map;
    }

    /**
     * 提取有效编号
     *
     * @param str
     * @param p
     * @return
     */
    public static String getData(String str, Pattern p) {
        String str2 = "";
        Matcher m = p.matcher(str);
        while (m.find()) {
            str2 = str2 + m.group();
        }
        return str2;
    }

    /**
     * 处理id相同的数据
     * 将有重复出现的ID保存
     *
     * @param MELedgerList
     * @return
     */
    public static Map HandleCellData(List<MiddleEastAll> MELedgerList) {
        Set set = new HashSet();//定义set来标识ID相同的数据
        List<String> list = new ArrayList<>();
        for (int i = 1; i < MELedgerList.size(); i++) {
            String id = MELedgerList.get(i - 1).getId();
            String id1 = MELedgerList.get(i).getId();
            if (id.equals(id1)) {
                if (set.add(id1)) {
                    list.add(id1);
                }
            }
        }
        //将ID重复的数据全部取出，分组保存
        Map<String, List<MiddleEastAll>> map = new HashMap<>();
        if (list.size() > 0) {
            for (String key : list) {
                //将所有的单元格合并数据，分组保存
                List<MiddleEastAll> middleEList = new ArrayList<>();
                for (MiddleEastAll middleEast : MELedgerList) {
                    if (key.equals(middleEast.getId())) {
                        middleEList.add(middleEast);
                    }
                }
                map.put(key, middleEList);
                System.out.println("有重复数据的ID值：" + key);
            }
        }
        return map;
    }

    /**
     * AR核销数据处理生成SAP表
     *
     * @param middleEastSAPList
     */
    public static void excelOutputSAP_AR(List<MiddleEastSAP> middleEastSAPList, String filePath) throws Exception {
        //创建表格
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        //定义第一个sheet页
        XSSFSheet xssfSheet = xssfWorkbook.createSheet("Sheet1");
        //第一个sheet页数据（生成台账表）
        Row row0 = xssfSheet.createRow(0);
        String[] headers = new String[]{"ID", "No", "Document date", "Type", "Company Code", "Posting Date", "Period", "Currency/Rate", "Document Number", "Translatn Date", "Reference", "Cross-CC No.", "Doc.Header Text", "Branch Number", "Trading Part.BA", "Postkey", "GL", "SGL Ind", "Amount", "Tax Code", "Business Area", "Cost Center", "Profit Center", "Value Date", "Due On", "Bline Date", "Issue Date", "Ext. No", "Bank/Acct No", "Disc Base", "Assignment", "Text", "Long Text", "Reason code"};
        for (int i = 0; i < headers.length; i++) {
            XSSFCell cell = (XSSFCell) row0.createCell(i);
            xssfSheet.setColumnWidth(i, 5000);
            XSSFRichTextString text = new XSSFRichTextString(headers[i]);
            cell.setCellValue(text);
        }
        int rowNum = 1;
        for (MiddleEastSAP MEastSap : middleEastSAPList) {
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
            if (!"".equals(amount) && amount != null) {
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
            rowNum++;
        }
        FileOutputStream fileOutputStream = new FileOutputStream(filePath);
        xssfWorkbook.write(fileOutputStream);
    }

    /**
     * AR核销数据处理生成F-51表
     *
     * @param middleEastSAPList
     */
    public static void excelOutput51_AR(List<MiddleEastSAP> middleEastSAPList, String filePath) throws Exception {
        //创建表格
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        //定义第一个sheet页
        XSSFSheet xssfSheet = xssfWorkbook.createSheet("Sheet1");
        //第一个sheet页数据（生成台账表）
        Row row0 = xssfSheet.createRow(0);
        String[] headers = new String[]{"ID", "No", "Document date", "Type", "Company Code", "Posting Date", "Period", "Currency/Rate", "Reference", "Cross-CC No.", "Doc.Header Text", "Branch Number", "Trading Part.BA", "Postkey", "GL", "SGL Ind", "Amount", "Tax Code", "Business Area", "Cost Center", "Profit Center", "Value Date", "Due On", "Bline Date", "Issue Date", "Ext. No", "Bank/Acct No", "Disc Base", "Assignment", "Text", "Long Text", "Reason code", "SAP No. for Other"};
        for (int i = 0; i < headers.length; i++) {
            XSSFCell cell = (XSSFCell) row0.createCell(i);
            xssfSheet.setColumnWidth(i, 5000);
            XSSFRichTextString text = new XSSFRichTextString(headers[i]);
            cell.setCellValue(text);
        }
        int rowNum = 1;
        for (MiddleEastSAP MEastSap : middleEastSAPList) {
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
            row.createCell(8).setCellValue(MEastSap.getReference());
            row.createCell(9).setCellValue(MEastSap.getCrossCCNo());
            row.createCell(10).setCellValue(MEastSap.getDocHeaderText());
            row.createCell(11).setCellValue(MEastSap.getBranchNumber());
            row.createCell(12).setCellValue(MEastSap.getTradingPartBA());
            row.createCell(13).setCellValue(MEastSap.getPostKey());
            row.createCell(14).setCellValue(MEastSap.getGl());
            row.createCell(15).setCellValue(MEastSap.getSglInd());
            String amount = MEastSap.getAmount();
            if (!"".equals(amount) && amount != null) {
                Pattern p = Pattern.compile("[^.0-9]");//提取有效数字
                amount = p.matcher(amount).replaceAll("").trim();
                row.createCell(16).setCellValue(convertData(amount));
            }
            row.createCell(17).setCellValue(MEastSap.getTaxCode());
            row.createCell(18).setCellValue(MEastSap.getBusinessArea());
            row.createCell(19).setCellValue(MEastSap.getCostCenter());
            row.createCell(20).setCellValue(MEastSap.getProfitCenter());
            row.createCell(21).setCellValue(MEastSap.getValueDate());
            row.createCell(22).setCellValue(MEastSap.getDuOn());
            row.createCell(23).setCellValue(MEastSap.getBlinDate());
            row.createCell(24).setCellValue(MEastSap.getIssueDate());
            row.createCell(25).setCellValue(MEastSap.getExtNo());
            row.createCell(26).setCellValue(MEastSap.getBankAcctNo());
            row.createCell(27).setCellValue(MEastSap.getDiscBase());
            row.createCell(28).setCellValue(MEastSap.getAssignment());
            row.createCell(29).setCellValue(MEastSap.getText());
            row.createCell(30).setCellValue(MEastSap.getLongText());
            row.createCell(31).setCellValue(MEastSap.getReasonCode());
            row.createCell(32).setCellValue(MEastSap.getSapforother());
            rowNum++;
        }
        FileOutputStream fileOutputStream = new FileOutputStream(filePath);
        xssfWorkbook.write(fileOutputStream);
    }
}