package HttpUtil;

import DataClean.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.util.*;
import java.util.regex.Pattern;

public class Russia_ARTwo extends BaseUtil {
    public static void main(String[] args) {
        String filePath = args[0];//水单路径
        String filePath1 = args[1];//101对照表
        String filePath_cus =args[2];//101cutomerlist对照表
        String filePath1_other =args[3];//101对照other表
        String filePath2 = args[4];//102对照表
        String filePath3 = args[5];//103对照表
//
//        String filePath = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\俄罗斯AR2\\201907081349\\0705-101.xlsx";
//        String filePath1 = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\俄罗斯AR2\\101匹配表.xlsx";
//        String filePath_cus = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\俄罗斯AR\\62F0customerlist.xlsx";//对照表
//        String filePath1_other = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\俄罗斯AR\\代办\\vendorlist_101.xlsx";
//        String filePath2 = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\俄罗斯AR2\\102扣款vendor_list匹配表.xlsx";
//        String filePath3 = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\俄罗斯AR2\\103账户客户编码匹配.xlsx";

        String code = filePath.substring(filePath.lastIndexOf("-") + 1, filePath.lastIndexOf("."));
        String filePath_SAP = filePath.substring(0, filePath.lastIndexOf("\\") + 1) + "FB01.xls";
        String ledger_Path = filePath.substring(0, filePath.lastIndexOf("\\") + 1) + "basicinfo_pay.xlsx";
        System.out.println("编号：" + code);
        try {
            List<RussiaAR> RussiaARList = null;
            switch (code) {
                case "101":
                    // 101 付款
                    List<RussiaLedger> LedList = getLedgerFunction(filePath);
                    if (LedList.size() > 0) {
                        RussiaARList = HandleData(LedList, ledger_Path, filePath1, filePath1_other, filePath_cus);
                    }
//                    for (RussiaAR russiaAR : RussiaARList) {
//                        if (russiaAR.getText()!=null){
//                            String text = russiaAR.getText();
//                            text = text.replace("Оплата по счету","Оп. сч.");
//                            text = text.replace("счет","сч.");
//                            russiaAR.setText(text);
//                        }
//                    }
                    break;
                case "102":
                    // 102 收款
                    List<RussiaLedger> LedList2 = getLedgerFunction2(filePath);
                    String ledger_Path1 = filePath.substring(0, filePath.lastIndexOf("\\") + 1) + "basicinfo_receive.xlsx";
                    if (LedList2.size() > 0) {
                        RussiaARList = HandleData2(LedList2, ledger_Path1);
                        for (RussiaAR russiaAR : RussiaARList) {
                            System.out.println("收款里面的：" + russiaAR.getNo());
                        }
                    } else {
                        RussiaARList = new ArrayList();
                    }
                    // 102 付款
                    List<RussiaLedger> LedList02 = getLedgerFunction(filePath);
                    if (LedList02.size() > 0) {
                        List<RussiaAR> RussiaARList2 = HandleData02(LedList02, ledger_Path, filePath2);
                        for (RussiaAR russiaAR : RussiaARList2) {
                            System.out.println("付款里面的：" + russiaAR.getNo());

                        }
                        if (RussiaARList2.size() > 0) {
                            System.out.println(RussiaARList2.size());
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
                excelOutput_FB01(RussiaARList, filePath_SAP);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 101编号 付款 数据处理
     */
    public static List HandleData(List<RussiaLedger> LedList, String ledger_Path, String banklist, String otherBankList, String filePath_cus) throws Exception {
        System.out.println("数据处理中！");
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
//            把title都大写化
            title = title.toUpperCase();
            if (title.indexOf(("Комиссия за валютный контроль").toUpperCase()) == 0 || (title.contains(title.replace("Комиссия за валютный контроль", "").trim().toUpperCase()) && title.contains("НДС на комиссию за валютный контроль".toUpperCase()))) {
                System.out.println("走的1");
                if ((title.contains(title.replace("Комиссия за валютный контроль", "").trim().toUpperCase()) && title.contains("НДС на комиссию за валютный контроль".toUpperCase()))) {
                    continue;
                } else {
                    String str = title.replace("Комиссия за валютный контроль".toUpperCase(), "").trim();
                    String amount = "";
                    RussiaLedger ledgers = null;
                    for (RussiaLedger ledger1 : LedList) {
                        String title1 = ledger1.getTitle().trim().toUpperCase();
                        if (title1.contains(str.toUpperCase()) && title1.contains("НДС на комиссию за валютный контроль".toUpperCase())) {
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
                        russiaAR1.setType("KZ");
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
                        russiaAR2.setType("KZ");
                        russiaAR2.setCompanyCode("62F0");
                        russiaAR2.setPostingDate(ledger.getDate());
                        russiaAR2.setPeriod(ledger.getDate().substring(0, i));
                        russiaAR2.setCurrencyRate("RUB");
                        russiaAR2.setPostkey("25");
                        russiaAR2.setGl("V999995830");
                        russiaAR2.setAmount("*");
                       if (cleanTitle(ledger.getTitle()).length() > 50) {
                        russiaAR2.setText(cleanTitle(ledger.getTitle()).substring(0, 50));
                    } else {
                        russiaAR2.setText(cleanTitle(ledger.getTitle()));
                    }
                        RussiaList.add(russiaAR2);
                    } else {
                        LedgerList.add(ledger);
                    }
                }
                continue;
            }
            if (title.indexOf("Комиссия за выбор банка-посредника согласно тарифам банка".toUpperCase()) == 0 || (title.contains(title.replace("Комиссия за выбор банка-посредника согласно тарифам банка", "").trim().toUpperCase()) && title.contains("НДС по комиссии за выбор банка-посредника".toUpperCase()))) {
                System.out.println("走的2");
                if ((title.contains(title.replace("Комиссия за выбор банка-посредника согласно тарифам банка".toUpperCase(), "").trim().toUpperCase()) && title.contains("НДС по комиссии за выбор банка-посредника".toUpperCase()))) {
                    System.out.println("过滤掉第二个");
                    continue;
                } else {
                    System.out.println("添加掉第一个");
                    String str = title.replace("Комиссия за выбор банка-посредника согласно тарифам банка".toUpperCase(), "").trim();
                    String amount = "";
                    RussiaLedger ledgers = null;
                    for (RussiaLedger ledger1 : LedList) {
                        String title1 = ledger1.getTitle().trim().toUpperCase();
                        if (title1.contains(str.toUpperCase()) && title1.contains("НДС по комиссии за выбор банка-посредника".toUpperCase())) {
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
                        russiaAR1.setType("KZ");
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
                        russiaAR2.setType("KZ");
                        russiaAR2.setCompanyCode("62F0");
                        russiaAR2.setPostingDate(ledger.getDate());
                        russiaAR2.setPeriod(ledger.getDate().substring(0, i));
                        russiaAR2.setCurrencyRate("RUB");
                        russiaAR2.setPostkey("25");
                        russiaAR2.setGl("V999995830");
                        russiaAR2.setAmount("*");
                       if (cleanTitle(ledger.getTitle()).length() > 50) {
                        russiaAR2.setText(cleanTitle(ledger.getTitle()).substring(0, 50));
                    } else {
                        russiaAR2.setText(cleanTitle(ledger.getTitle()));
                    }
                        RussiaList.add(russiaAR2);
                    } else {
                        LedgerList.add(ledger);
                    }
                }
                continue;
            }
            if (title.indexOf("Комиссия".toUpperCase()) == 0 && !title.contains("валютный контроль".toUpperCase()) && !title.contains("отправку документов".toUpperCase())) {
                System.out.println("走的3");
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
                if (cleanTitle(ledger.getTitle()).length() > 50) {
                        russiaAR1.setText(cleanTitle(ledger.getTitle()).substring(0, 50));
                    } 
                    else {
                        russiaAR1.setText(cleanTitle(ledger.getTitle()));
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
                if (cleanTitle(ledger.getTitle()).length() > 50) {
                    russiaAR2.setText(cleanTitle(ledger.getTitle()).substring(0, 50));
                }
                else {
                    russiaAR2.setText(cleanTitle(ledger.getTitle()));
                }
                RussiaList.add(russiaAR2);
            }
            if (title.contains("Перевод собственных средств. НДС не обл".toUpperCase()) || title.contains("Перевод собственных денежных средств".toUpperCase())) {
                System.out.println("走的4");
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
                if (cleanTitle(ledger.getTitle()).length() > 50) {
                        russiaAR1.setText(cleanTitle(ledger.getTitle()).substring(0, 50));
                    } 
                    else {
                        russiaAR1.setText(cleanTitle(ledger.getTitle()));
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
                if (cleanTitle(ledger.getTitle()).length() > 50) {
                    russiaAR2.setText(cleanTitle(ledger.getTitle()).substring(0, 50));
                }
                else {
                    russiaAR2.setText(cleanTitle(ledger.getTitle()));
                }
                russiaAR2.setReasonCode("100");
                RussiaList.add(russiaAR2);
            }
            if (title.contains("Размещение депозита".toUpperCase())) {
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
                if (cleanTitle(ledger.getTitle()).length() > 50) {
                        russiaAR1.setText(cleanTitle(ledger.getTitle()).substring(0, 50));
                    } 
                    else {
                        russiaAR1.setText(cleanTitle(ledger.getTitle()));
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
                continue;
            }
            if (title.contains("Покупка".toUpperCase())) {
                System.out.println("走的5");
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
                if (cleanTitle(ledger.getTitle()).length() > 50) {
                        russiaAR1.setText(cleanTitle(ledger.getTitle()).substring(0, 50));
                    } 
                    else {
                        russiaAR1.setText(cleanTitle(ledger.getTitle()));
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
                if (cleanTitle(ledger.getTitle()).length() > 50) {
                    russiaAR2.setText(cleanTitle(ledger.getTitle()).substring(0, 50));
                }
                else {
                    russiaAR2.setText(cleanTitle(ledger.getTitle()));
                }
                russiaAR2.setReasonCode("100");
                RussiaList.add(russiaAR2);
                continue;
            }
            if (title.contains("НДФЛ".toUpperCase())) {
                System.out.println("走的6");
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
                if (cleanTitle(ledger.getTitle()).length() > 50) {
                        russiaAR1.setText(cleanTitle(ledger.getTitle()).substring(0, 50));
                    } 
                    else {
                        russiaAR1.setText(cleanTitle(ledger.getTitle()));
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
                if (cleanTitle(ledger.getTitle()).length() > 50) {
                    russiaAR2.setText(cleanTitle(ledger.getTitle()).substring(0, 50));
                }
                else {
                    russiaAR2.setText(cleanTitle(ledger.getTitle()));
                }
                russiaAR2.setReasonCode("");
                RussiaList.add(russiaAR2);
                continue;
            }
            if (title.contains("Платеж на страхование от НС".toUpperCase())) {
                System.out.println("走的7");
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
                if (cleanTitle(ledger.getTitle()).length() > 50) {
                        russiaAR1.setText(cleanTitle(ledger.getTitle()).substring(0, 50));
                    } 
                    else {
                        russiaAR1.setText(cleanTitle(ledger.getTitle()));
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
                if (cleanTitle(ledger.getTitle()).length() > 50) {
                    russiaAR2.setText(cleanTitle(ledger.getTitle()).substring(0, 50));
                }
                else {
                    russiaAR2.setText(cleanTitle(ledger.getTitle()));
                }
                russiaAR2.setReasonCode("");
                RussiaList.add(russiaAR2);
                continue;
            }
            if (title.contains("Авансовые платежи зачисляемый ы ФСС".toUpperCase())) {
                System.out.println("走的9");
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
                if (cleanTitle(ledger.getTitle()).length() > 50) {
                        russiaAR1.setText(cleanTitle(ledger.getTitle()).substring(0, 50));
                    } 
                    else {
                        russiaAR1.setText(cleanTitle(ledger.getTitle()));
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
                if (cleanTitle(ledger.getTitle()).length() > 50) {
                    russiaAR2.setText(cleanTitle(ledger.getTitle()).substring(0, 50));
                }
                else {
                    russiaAR2.setText(cleanTitle(ledger.getTitle()));
                }
                russiaAR2.setReasonCode("");
                RussiaList.add(russiaAR2);
                continue;
            }
            if (title.contains("Страховый взносы на ОМС зачисляемые в бюджет  ФФОМС".toUpperCase())) {
                System.out.println("走的10");
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
                if (cleanTitle(ledger.getTitle()).length() > 50) {
                        russiaAR1.setText(cleanTitle(ledger.getTitle()).substring(0, 50));
                    } 
                    else {
                        russiaAR1.setText(cleanTitle(ledger.getTitle()));
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
                if (cleanTitle(ledger.getTitle()).length() > 50) {
                    russiaAR2.setText(cleanTitle(ledger.getTitle()).substring(0, 50));
                }
                else {
                    russiaAR2.setText(cleanTitle(ledger.getTitle()));
                }
                russiaAR2.setReasonCode("");
                RussiaList.add(russiaAR2);
                continue;
            }
            if (title.contains("Страх. Взносы на выплату строховой части трудовой пенси".toUpperCase())) {
                System.out.println("走的11");
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
                if (cleanTitle(ledger.getTitle()).length() > 50) {
                        russiaAR1.setText(cleanTitle(ledger.getTitle()).substring(0, 50));
                    } 
                    else {
                        russiaAR1.setText(cleanTitle(ledger.getTitle()));
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
                if (cleanTitle(ledger.getTitle()).length() > 50) {
                    russiaAR2.setText(cleanTitle(ledger.getTitle()).substring(0, 50));
                }
                else {
                    russiaAR2.setText(cleanTitle(ledger.getTitle()));
                }
                russiaAR2.setReasonCode("");
                RussiaList.add(russiaAR2);
                continue;
            }
//            退款title中包含претензии
            if (title.contains("претензии".toUpperCase())) {
                System.out.println("走的12");
                if (set.add(ledger.getNo())) {
                    ledger.setReasonCode("172");
                    ledger.setVendorCode(getrefundAccount(ledger.getTax().substring(1, ledger.getTax().length() - 1), filePath_cus));
                    LedgerList.add(ledger);
                } else {
                    int f = LedgerList.indexOf(ledger);
                    LedgerList.remove(f);
                    ledger.setReasonCode("172");
                    ledger.setVendorCode(getrefundAccount(ledger.getTax().substring(1, ledger.getTax().length() - 1), filePath_cus));
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
                if (cleanTitle(ledger.getTitle()).length() > 50) {
                        russiaAR1.setText(cleanTitle(ledger.getTitle()).substring(0, 50));
                    } 
                    else {
                        russiaAR1.setText(cleanTitle(ledger.getTitle()));
                    }
                russiaAR1.setReasonCode("172");
                RussiaList.add(russiaAR1);

                russiaAR2.setNo(ledger.getNo());
                russiaAR2.setDocumentDate(ledger.getDate());
                russiaAR2.setType("KZ");
                russiaAR2.setCompanyCode("62F0");
                russiaAR2.setPostingDate(ledger.getDate());
                russiaAR2.setPeriod(ledger.getDate().substring(0, i));
                russiaAR2.setCurrencyRate("RUB");
                russiaAR2.setSglInd("");
                russiaAR2.setPostkey("05");
                russiaAR2.setGl(getrefundAccount(ledger.getTax().substring(1, ledger.getTax().length() - 1), filePath_cus));
                russiaAR2.setAmount("*");
                russiaAR2.setDueOn(ledger.getDate());
                if (cleanTitle(ledger.getTitle()).length() > 50) {
                    russiaAR2.setText(cleanTitle(ledger.getTitle()).substring(0, 50));
                }
                else {
                    russiaAR2.setText(cleanTitle(ledger.getTitle()));
                }
                russiaAR2.setReasonCode("");
                RussiaList.add(russiaAR2);
                continue;
            }
//            其他情况Vendor依据customer内容通过匹配表匹配
            if (!(title.indexOf("Комиссия за валютный контроль".toUpperCase()) == 0) && !(title.indexOf("Комиссия за выбор банка-посредника согласно тарифам банка".toUpperCase()) == 0) && !(title.indexOf("Комиссия".toUpperCase()) == 0 && !title.contains("валютный контроль".toUpperCase()) && !title.contains("отправку документов".toUpperCase())) && !(title.contains("Перевод собственных средств. НДС не обл".toUpperCase()) || title.contains("Перевод собственных денежных средств".toUpperCase())) && !(title.contains("Размещение депозита".toUpperCase())) && !(title.contains("Покупка".toUpperCase())) && !(title.contains("НДФЛ".toUpperCase())) && !(title.contains("Платеж на страхование от НС".toUpperCase())) && !(title.contains("Авансовые платежи зачисляемый ы ФСС".toUpperCase())) && !(title.contains("Страховый взносы на ОМС зачисляемые в бюджет  ФФОМС".toUpperCase())) && !(title.contains("Страх. Взносы на выплату строховой части трудовой пенси".toUpperCase())) && !(title.contains("претензии".toUpperCase()))) {
                System.out.println("走的13");
//                台账
                if (set.add(ledger.getNo())) {
                    if (getOtherAccount(ledger.getVendor(), otherBankList).size() > 0) {
                        ledger.setReasonCode(getOtherAccount(ledger.getVendor(), otherBankList).get(1).toString());
                        ledger.setVendorCode(getOtherAccount(ledger.getVendor(), otherBankList).get(0).toString());
                    } else {
                        ledger.setReasonCode("");
                        ledger.setVendorCode("");
                    }
                    LedgerList.add(ledger);
                } else {
                    int f = LedgerList.indexOf(ledger);
                    LedgerList.remove(f);
                    if (getOtherAccount(ledger.getVendor(), otherBankList).size() > 0) {
                        ledger.setReasonCode(getOtherAccount(ledger.getVendor(), otherBankList).get(1).toString());
                        ledger.setVendorCode(getOtherAccount(ledger.getVendor(), otherBankList).get(0).toString());
                    } else {
                        ledger.setReasonCode("");
                        ledger.setVendorCode("");
                    }
                    LedgerList.add(ledger);
                }
                if (!(title.indexOf("НДС".toUpperCase()) == 0 || title.indexOf("Комиссия".toUpperCase()) == 0)) {
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
                    if (cleanTitle(ledger.getTitle()).length() > 50) {
                        russiaAR1.setText(cleanTitle(ledger.getTitle()).substring(0, 50));
                    } 
                    else {
                        russiaAR1.setText(cleanTitle(ledger.getTitle()));
                    }
                    if (getOtherAccount(ledger.getVendor(), otherBankList).size() > 0) {
                        russiaAR1.setReasonCode(getOtherAccount(ledger.getVendor(), otherBankList).get(1).toString());
                    } else {
                        russiaAR1.setReasonCode("");
                    }
                    if (!("".equals(russiaAR1.getGl())) && !("".equals(russiaAR1.getReasonCode()))) {
                        RussiaList.add(russiaAR1);
                    }
                    russiaAR2.setNo(ledger.getNo());
                    russiaAR2.setDocumentDate(ledger.getDate());
                    russiaAR2.setType("KZ");
                    russiaAR2.setCompanyCode("62F0");
                    russiaAR2.setPostingDate(ledger.getDate());
                    russiaAR2.setPeriod(ledger.getDate().substring(0, i));
                    russiaAR2.setCurrencyRate("RUB");
                    russiaAR2.setSglInd("");
                    russiaAR2.setPostkey("25");
                    if (getOtherAccount(ledger.getVendor(), otherBankList).size() > 0) {
                        russiaAR2.setGl(getOtherAccount(ledger.getVendor(), otherBankList).get(0).toString());
                    } else {
                        russiaAR2.setGl("");
                    }
                    russiaAR2.setAmount("*");
                    russiaAR2.setDueOn(ledger.getDate());
                    if (cleanTitle(ledger.getTitle()).length() > 50) {
                        russiaAR2.setText(cleanTitle(ledger.getTitle()).substring(0, 50));
                    } else {
                        russiaAR2.setText(cleanTitle(ledger.getTitle()));
                    }
                    russiaAR2.setReasonCode("");

                    if (!(russiaAR2.getGl().equals(""))) {
                        RussiaList.add(russiaAR2);
                    }
                }
                continue;
            }
//

            if (set.add(ledger.getNo())) {
                LedgerList.add(ledger);
            }
        }
        System.out.println("处理之后的台账：" + LedgerList.size());
        excelOutput_ledger(LedgerList, ledger_Path);
        return RussiaList;
    }

    private static String cleanTitle(String title) {
//        if (title.length() > 50) {
        title = title.substring(1);
        title = title.replaceAll("Оплата по счету", "Оп. сч");
        title = title.replaceAll("счет", "сч");
//        }
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
//            System.out.println("输出VendorCode:" + ledger.getVendorCode());
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
            String tax = "";
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
                if (row.getCell(8) != null) {
                    row.getCell(8).setCellType(Cell.CELL_TYPE_STRING);
                    tax = row.getCell(8).getStringCellValue();
                }
                if (row.getCell(9) != null) {
                    row.getCell(9).setCellType(Cell.CELL_TYPE_STRING);
                    taxCode = clearData(row.getCell(9).getStringCellValue());
                }
                String uuid = UUID.randomUUID().toString().replaceAll("-", "");
                ledger.setNo(uuid);
                ledger.setAmount(amount.replace(" ", ""));
                ledger.setDate(date);
                ledger.setTax(tax);
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
}
