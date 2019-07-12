package HttpUtil;

import DataClean.BaseUtil;
import DataClean.DataDO;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 新加坡生成ASP表业务处理
 */
public class ETLCombine extends BaseUtil {
    public static void main(String[] args)throws Exception{
//        String companycode = "62T0-HQ";
//        String ETL_path = "F:\\test\\ETLfile.xls";
//        String SAP_Path="F:\\test\\SAP_file.xls";
        String companycode=args[0];
        String ETL_path =args[1];
        String SAP_Path=args[2];
        File EIL_file = new File(ETL_path);
        String CompanyCode = "";
        Date OrderShipmentDate = new Date();
        Date ActualShipmentDate = new Date();
        String OrderShipmentDate_final = "";
        String ActualShipmentDate_final = "";
        String InvoiceReferenceNumber = "";
        String InvoiceReferenceNumber2 = "";
        String POShortText = "";
        String PurchaseOrderNumber = "";
        String Amount = "";
        String Quantity = "";
        String TaxAmount="";
        String TotalAmount = "";
        String UnitPrice = "";
        String Currency = "";
        String SONumber = "";
        String GoodsDescription = "";
        String InvoiceDate_final = "";
        String InvoiceDate = "";
        String OCRStatus = "";
        String FileName = "";
        try {
            //获取表格数据
            FileInputStream EIL_file_IO = new FileInputStream(EIL_file);
            Workbook xssfSheets_ETL = WorkbookFactory.create(EIL_file_IO);
            Sheet sheet_ETL = xssfSheets_ETL.getSheetAt(0);

            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("MM/dd/yyyy");
            List<DataDO> dataDO_ETL = new ArrayList<>();//用来存储所有数据
            ETLCombine etlCombine = new ETLCombine();
            //遍历表格数据
            for (int r = 1; r <= sheet_ETL.getLastRowNum(); r++) {
                Row row = sheet_ETL.getRow(r);
                DataDO dataDO = new DataDO();
                if(row.getCell(0) != null){
                    row.getCell(0).setCellType(Cell.CELL_TYPE_STRING);
                    FileName = row.getCell(0).getStringCellValue();
                }else {
                    FileName = "";
                }
                if(row.getCell(2) != null){
                    row.getCell(2).setCellType(Cell.CELL_TYPE_STRING);
                    CompanyCode = row.getCell(2).getStringCellValue();
                }else {
                    CompanyCode = "";
                }
                if(row.getCell(3) != null){
                    row.getCell(3).setCellType(Cell.CELL_TYPE_STRING);
                    Amount = row.getCell(3).getStringCellValue();
                }else {
                    Amount = "";
                }
                if(row.getCell(4) != null){
                    row.getCell(4).setCellType(Cell.CELL_TYPE_STRING);
                    InvoiceDate = row.getCell(4).getStringCellValue();
                }else {
                    InvoiceDate = "";
                }
                if(row.getCell(5) != null){
                    row.getCell(5).setCellType(Cell.CELL_TYPE_STRING);
                    InvoiceReferenceNumber = row.getCell(5).getStringCellValue();
                }
                else {
                    InvoiceReferenceNumber = "";
                }
                if(row.getCell(6) != null){
                    row.getCell(6).setCellType(Cell.CELL_TYPE_STRING);
                    InvoiceReferenceNumber2 = row.getCell(6).getStringCellValue();
                } else {
                    InvoiceReferenceNumber2 = "";
                }
                if (row.getCell(7) != null){
                    row.getCell(7).setCellType(Cell.CELL_TYPE_STRING);
                    POShortText = row.getCell(7).getStringCellValue();
                }else {
                    POShortText = "";
                }
                if (row.getCell(8) != null){
                    row.getCell(8).setCellType(Cell.CELL_TYPE_STRING);
                    PurchaseOrderNumber = row.getCell(8).getStringCellValue();
                }else {
                    PurchaseOrderNumber = "";
                }
                if (row.getCell(9) != null){
                    row.getCell(9).setCellType(Cell.CELL_TYPE_STRING);
                    Quantity = row.getCell(9).getStringCellValue();
                }else {
                    Quantity = "";
                }
                if (row.getCell(10) != null ){
                    row.getCell(10).setCellType(Cell.CELL_TYPE_STRING);
                    TaxAmount = row.getCell(10).getStringCellValue();
                } else {
                    TaxAmount = "";
                }
                if (row.getCell(11) != null){
                    row.getCell(11).setCellType(Cell.CELL_TYPE_STRING);
                    TotalAmount = row.getCell(11).getStringCellValue();
                }else {
                    TotalAmount = "";
                }
                if (row.getCell(12) != null){
                    row.getCell(12).setCellType(Cell.CELL_TYPE_STRING);
                    UnitPrice = row.getCell(12).getStringCellValue();
                }else {
                    UnitPrice = "";
                }
                if (row.getCell(13) != null){
                    row.getCell(13).setCellType(Cell.CELL_TYPE_STRING);
                    Currency = row.getCell(13).getStringCellValue();
                }else {
                    Currency = "";
                }
                if (row.getCell(14) != null){
                    row.getCell(14).setCellType(Cell.CELL_TYPE_STRING);
                    SONumber =  row.getCell(14).getStringCellValue();
                } else {
                    SONumber = "";
                }
                if (row.getCell(15) != null){
                    row.getCell(15).setCellType(Cell.CELL_TYPE_STRING);
                    GoodsDescription = row.getCell(15).getStringCellValue();
                }else {
                    GoodsDescription = "";
                }
                if (row.getCell(16) != null){
                    row.getCell(16).setCellType(Cell.CELL_TYPE_STRING);
                    OCRStatus = row.getCell(16).getStringCellValue();
                }else {
                    OCRStatus = "";
                }
                if (row.getCell(17) != null){
                    OrderShipmentDate = row.getCell(17).getDateCellValue();
                    OrderShipmentDate_final = simpleDateFormat.format(OrderShipmentDate);
                } else {
                    OrderShipmentDate_final = "";
                }
                if (row.getCell(18) != null){
                    ActualShipmentDate = row.getCell(18).getDateCellValue();
                    ActualShipmentDate_final = simpleDateFormat.format(ActualShipmentDate);
                } else {
                    ActualShipmentDate_final = "";
                }
                dataDO.setFilepath(FileName);
                dataDO.setCompanyCode(CompanyCode);
                dataDO.setAmount(convertData(Amount));
                dataDO.setInvoicedate(ActualShipmentDate_final);
                dataDO.setInvoiceReferenceNumber(InvoiceReferenceNumber);
                dataDO.setInvoiceReferenceNumber2(InvoiceReferenceNumber2);
                dataDO.setPoshorttext(POShortText);
                dataDO.setPurchaseOrderNumber(PurchaseOrderNumber);
                dataDO.setQuantity(Quantity);
                dataDO.setTaxAmount(TaxAmount);
                dataDO.setTotalAmount(convertData(TotalAmount));
                dataDO.setUnitPrice(convertData(UnitPrice));
                dataDO.setSOnumber(SONumber);
                dataDO.setCurrency(Currency);
                dataDO.setGoodDescription(GoodsDescription);
                dataDO.setOCRStatus(OCRStatus);
                dataDO.setOrderShipmentDate(OrderShipmentDate_final);
                dataDO.setActualShipmentDate(ActualShipmentDate_final);
                dataDO_ETL.add(dataDO);
            }
            List<DataDO> dataDO_Ex = new ArrayList<>();//存储字段缺失数据
            List<DataDO> dataDO = new ArrayList<>();//存储正常数据
            for (DataDO d : dataDO_ETL){//遍历所有处理好的数据
                String status = "";
                if (!"OCR Data Incompleted".equals(d.getOCRStatus())){
                    if (!"".equals(d.getInvoiceReferenceNumber()) && !"".equals(d.getAmount()) && !"".equals(d.getQuantity()) && !"".equals(d.getUnitPrice()) && !"".equals(d.getTotalAmount()) && !"".equals(d.getPurchaseOrderNumber())){
                        if(!"".equals(d.getOrderShipmentDate()) && !"".equals(d.getActualShipmentDate())){
                            dataDO.add(d);
                        } else {
                            d.setOCRStatus("HROIS Cannot Detect");
                            dataDO_Ex.add(d) ;//字段缺失数据
                        }
                    }else {
                        if ("".equals(d.getInvoiceReferenceNumber())){
                            status = status + "InvoiceReferenceNumber Cannot Detect!";
                        }
                        if ("".equals(d.getAmount())){
                            status = status + "Amount Cannot Detect!";
                        }
                        if ("".equals("".equals(d.getQuantity()))){
                            status = status + "Quantity Cannot Detect!";
                        }
                        if ("".equals(d.getUnitPrice())){
                            status = status + "UnitPrice Cannot Detect!";
                        }
                        if ("".equals(d.getTotalAmount())){
                            status = status + "TotalAmount Cannot Detect!";
                        }
                        if ("".equals(d.getPurchaseOrderNumber())){
                            status = status + "PurchaseOrderNumber Cannot Detect!";
                        }
                        d.setOCRStatus(status);
                        dataDO_Ex.add(d) ;//字段缺失数据
                    }
                }else {
                    dataDO_Ex.add(d) ;//字段缺失数据
                }
            }
            if(dataDO.size() > 0){
                // 生成SAP表
                List<DataDO> dataDO_SAP=  etlCombine.invoiceVerification(dataDO);
                etlCombine.excelOutput(dataDO_SAP,SAP_Path);//正常数据
            }
            EIL_file.delete();//删除文件
            etlCombine.excelOutput_Exception_ETL(dataDO_ETL,ETL_path);//生成日志表
                //etlCombine.excelOutput_Exception_ETL(dataDO_ETL,SAP_Path_EX);//日志表
        } catch (Exception e) {
            e.printStackTrace();
        }

    }
    public void excelOutput(List<DataDO> dataDOS,String filename){
        XSSFWorkbook xssfWorkbook =new XSSFWorkbook();
        XSSFSheet xssfSheet=xssfWorkbook.createSheet("Sheet1");
        String[] headers=new String[]{};
        headers = new String[]{"FileName","CompanyCode", "Amount", "InvoiceDate", "InvoiceReferenceNumber","InvoiceReferenceNumber2", "POShortText", "PurchaseOrderNumber", "Quantity", "TaxAmount", "TotalAmount", "UnitPrice","Currency","SONumber","GoodsDescription","OCRStatus","PostingDate", "TaxCode", "Text", "BaselineDate", "ExchangeRate", "PaymentBlock", "Assignment","HeaderText"};
        Row row0=xssfSheet.createRow(0);
        for(int i=0;i<headers.length;i++){
            XSSFCell cell = (XSSFCell) row0.createCell(i);
            XSSFRichTextString text = new XSSFRichTextString(headers[i]);
            cell.setCellValue(text);
        }
        int rowNum = 1;
        for (DataDO dataDO:dataDOS){
            XSSFRow row1 = xssfSheet.createRow(rowNum);
                row1.createCell(0).setCellType(Cell.CELL_TYPE_STRING);
                row1.createCell(1).setCellType(Cell.CELL_TYPE_STRING);
                row1.createCell(2).setCellType(Cell.CELL_TYPE_STRING);
                row1.createCell(3).setCellType(Cell.CELL_TYPE_STRING);
                row1.createCell(4).setCellType(Cell.CELL_TYPE_STRING);
                row1.createCell(5).setCellType(Cell.CELL_TYPE_STRING);
                row1.createCell(6).setCellType(Cell.CELL_TYPE_STRING);
                row1.createCell(7).setCellType(Cell.CELL_TYPE_STRING);
                row1.createCell(8).setCellType(Cell.CELL_TYPE_STRING);
                row1.createCell(9).setCellType(Cell.CELL_TYPE_STRING);
                row1.createCell(10).setCellType(Cell.CELL_TYPE_STRING);
                row1.createCell(11).setCellType(Cell.CELL_TYPE_STRING);
                row1.createCell(12).setCellType(Cell.CELL_TYPE_STRING);
                row1.createCell(13).setCellType(Cell.CELL_TYPE_STRING);
                row1.createCell(14).setCellType(Cell.CELL_TYPE_STRING);
                row1.createCell(15).setCellType(Cell.CELL_TYPE_STRING);
                row1.createCell(16).setCellType(Cell.CELL_TYPE_STRING);
                row1.createCell(17).setCellType(Cell.CELL_TYPE_STRING);
                row1.createCell(18).setCellType(Cell.CELL_TYPE_STRING);
                row1.createCell(19).setCellType(Cell.CELL_TYPE_STRING);
                row1.createCell(20).setCellType(Cell.CELL_TYPE_STRING);
                row1.createCell(21).setCellType(Cell.CELL_TYPE_STRING);
                row1.createCell(22).setCellType(Cell.CELL_TYPE_STRING);
                row1.createCell(23).setCellType(Cell.CELL_TYPE_STRING);
                row1.getCell(0).setCellValue(dataDO.getFilepath());
                row1.getCell(1).setCellValue(dataDO.getCompanyCode());
                String companyCode = "";
                int count = 0;
                Pattern p = Pattern.compile("-");
                Matcher m = p.matcher(dataDO.getCompanyCode());
                while (m.find()) {
                    count++;
                }
                if (count == 1) {
                    companyCode = dataDO.getCompanyCode();
                }else {
                    companyCode = dataDO.getCompanyCode().substring(0,dataDO.getCompanyCode().lastIndexOf("-"));
                }
                row1.getCell(2).setCellValue(dataDO.getAmount());
                row1.getCell(2).setCellType(Cell.CELL_TYPE_STRING);
                row1.getCell(3).setCellValue(dataDO.getActualShipmentDate());
                row1.getCell(4).setCellValue(dataDO.getInvoiceReferenceNumber());
                row1.getCell(5).setCellValue(dataDO.getInvoiceReferenceNumber2());
                row1.getCell(6).setCellValue(dataDO.getPoshorttext());
                row1.getCell(7).setCellValue(dataDO.getPurchaseOrderNumber());
                row1.getCell(8).setCellValue(dataDO.getQuantity()+"".toString());
                row1.getCell(9).setCellValue(dataDO.getTaxAmount()+"".toString());
                row1.getCell(10).setCellValue(dataDO.getTotalAmount()+"".toString());
                row1.getCell(11).setCellValue(dataDO.getUnitPrice()+"".toString());
                row1.getCell(12).setCellValue(dataDO.getCurrency());
                row1.getCell(13).setCellValue(dataDO.getSOnumber());
                row1.getCell(14).setCellValue(dataDO.getGoodDescription());
                row1.getCell(15).setCellValue("OK");
                row1.getCell(16).setCellValue(dataDO.getPostingDate());
                row1.getCell(17).setCellValue(dataDO.getTaxCode());
                row1.getCell(18).setCellValue(dataDO.getPurchaseOrderNumber());
                String code  = companyCode.substring(0,companyCode.indexOf("-"));
            switch (code){
                case "62G0" :
                    System.out.println("再生成baselinedate的时候法国运行到这里号为");
                    row1.getCell(19).setCellValue(dataDO.getActualShipmentDate());
                    System.out.println("已经生成完baselinedatet的时候号为"+ row1.getCell(19)+"\n");
                    break;
                case "62H0" :
                    row1.getCell(19).setCellValue(dataDO.getActualShipmentDate());
                    break;
                case "62T0" :
                    row1.getCell(19).setCellValue(dataDO.getActualShipmentDate());
                    break;
                case "6280" :
                    row1.getCell(19).setCellValue(dataDO.getActualShipmentDate());
                    break;
                case "6200" :
                    row1.getCell(19).setCellValue(dataDO.getActualShipmentDate());
                    break;
                case "62S0" :
                    row1.getCell(19).setCellValue(dataDO.getActualShipmentDate());
                    break;
                case "6400" :
                    row1.getCell(19).setCellValue("");
                    break;
                case "6550" :
                    row1.getCell(19).setCellValue(dataDO.getActualShipmentDate());
                    break;
                case "6620" :
                    row1.getCell(19).setCellValue(dataDO.getActualShipmentDate());
                    break;
                default:
                    row1.getCell(19).setCellValue(dataDO.getInvoicedate());
            }
            row1.getCell(20).setCellValue(dataDO.getExchangeRate());
            row1.getCell(21).setCellValue(dataDO.getPaymentBlock());
            switch (code){
                case "62G0":
                    row1.getCell(22).setCellValue(dataDO.getPurchaseOrderNumber());
                    break;
                case "62H0":
                    row1.getCell(22).setCellValue(dataDO.getPurchaseOrderNumber());
                    break;
                case "62T0":
                    row1.getCell(22).setCellValue(dataDO.getPurchaseOrderNumber());
                    break;
                case "6280":
                    row1.getCell(22).setCellValue(dataDO.getPurchaseOrderNumber());
                    break;
                case "6200":
                    row1.getCell(22).setCellValue(dataDO.getPurchaseOrderNumber());
                    break;
                case "62S0":
                    row1.getCell(22).setCellValue(dataDO.getPurchaseOrderNumber());
                    break;
                case "6400":
                    row1.getCell(22).setCellValue(dataDO.getPurchaseOrderNumber());
                    break;
                case "6550":
                    row1.getCell(22).setCellValue(dataDO.getPurchaseOrderNumber());
                    break;
                case "65G0":
                    row1.getCell(22).setCellValue(dataDO.getPurchaseOrderNumber());
                    break;
                case "62F0-HQ":
                    row1.getCell(22).setCellValue(dataDO.getPurchaseOrderNumber());
                    break;
            }
            switch (code){
                case "62T0":
                    row1.getCell(23).setCellValue(dataDO.getPurchaseOrderNumber());
                    break;
                case "6200":
                    row1.getCell(23).setCellValue(dataDO.getPurchaseOrderNumber());
                    break;
                case "62S0":
                    row1.getCell(23).setCellValue(dataDO.getPurchaseOrderNumber());
                    break;
                case "6400":
                    row1.getCell(23).setCellValue(dataDO.getPurchaseOrderNumber());
                    break;
                case "6550":
                    row1.getCell(23).setCellValue(dataDO.getPurchaseOrderNumber());
                    break;
                case "65G0":
                    String folderNo = dataDO.getCompanyCode().substring(dataDO.getCompanyCode().lastIndexOf("-")+1);
                    row1.getCell(23).setCellValue(folderNo + "/" + dataDO.getInvoiceReferenceNumber() +"/" +dataDO.getPurchaseOrderNumber());
                    break;
            }
            rowNum++;
         }
        try {
            FileOutputStream fileOutputStream=new FileOutputStream(filename);
            try {
                xssfWorkbook.write(fileOutputStream);
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }
    public List<DataDO> invoiceVerification(List<DataDO> dataDOS){
        DateFormat dFormat = new SimpleDateFormat("MM/dd/yyyy");
        String formatDate = dFormat.format(new Date());
        List<DataDO> dataDOS1 = new ArrayList<>();
        if (dataDOS != null){
            for (DataDO dataDO:dataDOS){
                String companyCode = "";
                int count = 0;
                Pattern p = Pattern.compile("-");
                Matcher m = p.matcher(dataDO.getCompanyCode());
                while (m.find()) {
                    count++;
                }
                if (count == 1) {
                    companyCode = dataDO.getCompanyCode();
                }else {
                    companyCode = dataDO.getCompanyCode().substring(0,dataDO.getCompanyCode().lastIndexOf("-"));
                }
                String invoicedate=dataDO.getInvoicedate().substring(0,2);
                Calendar calendar = Calendar.getInstance();
                calendar.setTime(new Date());
                String month = formatDate.substring(0,2);
                String year = String.valueOf(calendar.get(Calendar.YEAR));
                String previousfirstday = invoicedate+"/01"+"/"+year;
                String nowfirstday=month+"/01/"+year;
                String code = companyCode.substring(0,companyCode.indexOf("-"));
                switch (code){
                    case "6550":
                        dataDO.setTaxCode("OP (Purchase in out of scope (e.g. overseas purchase))");
                        dataDO.setPostingDate(formatDate);
                        break;
                    case "6560":
                        dataDO.setTaxCode("V0 (Input VAT Receivable 0%)");
                        dataDO.setPostingDate(formatDate);
                        break;
                    case "65G0":
                        dataDO.setTaxCode("I0 (Input Tax Zero Rated)");
                        dataDO.setPostingDate(formatDate);
                        break;
                    case "6430":
                        dataDO.setTaxCode("X0 (no tax)");
                        if (invoicedate.equals(month)){
                            dataDO.setPostingDate(formatDate);
                        }else {
                            dataDO.setPostingDate(previousfirstday);
                        }
                        break;
                    case "62G0":
                        dataDO.setTaxCode("5E (I Acquisition Hors CE)");
                        dataDO.setPostingDate(formatDate);
                        break;
                    case "62H0":
                        dataDO.setTaxCode("B4 (AK - IMP - Purchase - 0%)");
                        dataDO.setPostingDate(formatDate);
                        break;
                    case "62T0":
                        dataDO.setTaxCode("3D (I Operaciones Extra Comunitarias 0%)");
                        dataDO.setPostingDate(formatDate);
                        break;
                    case "6280":
                        dataDO.setTaxCode("W0 (0% VAT Input Tax)");
                        dataDO.setPostingDate(formatDate);
                        break;
                    case "6200":
                        dataDO.setTaxCode("70 (FC IVA ART 7-bis, c1 DPR 633/72 -op non soggetta-)");
                        dataDO.setPostingDate(formatDate);
                        break;
                    case "62S0":
                        dataDO.setTaxCode("94 (Input VAT Exempt)");
                        dataDO.setPostingDate(formatDate);
                        break;
                    case "6620":
                        dataDO.setTaxCode("I0 (SALES TAX - INPUT(0%))");
                        dataDO.setPostingDate(formatDate);
                        break;
                    case "6400":
                        dataDO.setTaxCode("W9 ()");
                        dataDO.setPostingDate(formatDate);
                        break;
                    case "62F0":
                        dataDO.setTaxCode("P0 (0% Input Tax for Goods,Services)");
                        if (invoicedate.equals(month)){
                            dataDO.setPostingDate(dataDO.getInvoicedate());
                        }else {
                            dataDO.setPostingDate(nowfirstday);
                        }
                        break;
                }
                dataDO.setText("");
                dataDO.setBaselineDate("");
                dataDO.setExchangeRate("");
                dataDO.setBaselineDate("");
                dataDO.setPaymentBlock("");
                dataDO.setAssignment("");
                dataDO.setHeaderText("");
                dataDOS1.add(dataDO);
            }
        }
        return dataDOS1;
    }
    public void excelOutput_Exception_ETL(List<DataDO> dataDOS,String filename){
        XSSFWorkbook xssfWorkbook =new XSSFWorkbook();
        XSSFSheet xssfSheet=xssfWorkbook.createSheet("Sheet1");
        String[] headers=new String[]{};
        headers = new String[]{"FileName","DownloadStatus","CompanyCode", "Amount", "InvoiceDate", "InvoiceReferenceNumber","InvoiceReferenceNumber2", "POShortText", "PurchaseOrderNumber", "Quantity", "TaxAmount", "TotalAmount", "UnitPrice","Currency","SONumber","GoodsDescription","OCRStatus","OrderShipmentDate","ActualShipmentDate"};

        Row row0=xssfSheet.createRow(0);
        for(int i=0;i<headers.length;i++){
            XSSFCell cell = (XSSFCell) row0.createCell(i);
            XSSFRichTextString text = new XSSFRichTextString(headers[i]);
            cell.setCellValue(text);
        }
        int rowNum = 1;
        for (DataDO dataDO:dataDOS){
            XSSFRow row1 = xssfSheet.createRow(rowNum);
            row1.createCell(0).setCellValue(dataDO.getFilepath());
            row1.createCell(1).setCellValue("OK");
            row1.createCell(2).setCellValue(dataDO.getCompanyCode());
            row1.createCell(3).setCellValue(dataDO.getAmount());
            row1.createCell(4).setCellValue(dataDO.getActualShipmentDate());
            row1.createCell(5).setCellValue(dataDO.getInvoiceReferenceNumber());
            row1.createCell(6).setCellValue(dataDO.getInvoiceReferenceNumber2());
            row1.createCell(7).setCellValue(dataDO.getPoshorttext());
            row1.createCell(8).setCellValue(dataDO.getPurchaseOrderNumber());
            row1.createCell(9).setCellValue(dataDO.getQuantity());
            row1.createCell(10).setCellValue(dataDO.getTaxAmount());
            row1.createCell(11).setCellValue(dataDO.getTotalAmount());
            row1.createCell(12).setCellValue(dataDO.getUnitPrice());
            row1.createCell(13).setCellValue(dataDO.getCurrency());
            row1.createCell(14).setCellValue(dataDO.getSOnumber());
            row1.createCell(15).setCellValue(dataDO.getGoodDescription());
            row1.createCell(16).setCellValue(dataDO.getOCRStatus());
            row1.createCell(17).setCellValue(dataDO.getOrderShipmentDate());
            row1.createCell(18).setCellValue(dataDO.getActualShipmentDate());
            rowNum++;
        }
        try {
            FileOutputStream fileOutputStream=new FileOutputStream(filename);
            try {
                xssfWorkbook.write(fileOutputStream);
                fileOutputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }
}
