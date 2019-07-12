package HttpUtil;

import DataClean.BaseUtil;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.plaf.synth.Region;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Locale;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 表格拆分
 */
public class SplitTable extends BaseUtil {
    public static void main(String[] args) {
        String file_Path = "F:\\test\\collection\\2018 collection(12).xlsx";
        String fileName = "F:\\test\\collection\\Sheet_Excel" +".xls";
        try {
            XSSFWorkbook xssfWorkbook = splitExcel(file_Path);
            FileOutputStream fileOut = new FileOutputStream(fileName);
            xssfWorkbook.write(fileOut);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     *Collection表格拆分操作类
     * @param filePath
     * @return
     * @throws Exception
     */
    public static XSSFWorkbook splitExcel(String filePath) throws Exception{
            File EIL_file = new File(filePath);
            SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
            FileInputStream EIL_file_IO = new FileInputStream(EIL_file);
            Workbook xssfSheets_ETL =  WorkbookFactory.create(EIL_file_IO);
            //获取表格数据
            int index = 0;
            int num = 1;
            Sheet sheet_ETL = xssfSheets_ETL.getSheet("August 2018");
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
            for (int r = index; r <= sheet_ETL.getLastRowNum(); r++) {
                Row row = sheet_ETL.getRow(r); //获取每一行数据
                if (row != null){
                    if (row.getCell(0) != null || row.getCell(4) != null || row.getCell(5) != null){//判断获取到的每一行数据的第一列是否为null
                        String rowData = "";
                        if (row.getCell(0) != null){
                            row.getCell(0).setCellType(Cell.CELL_TYPE_STRING);
                            rowData = row.getCell(0).getStringCellValue();//获取每一行的第一列内容
                        }
                        if ("COLLECTION REPORT".equals(rowData)){//当获取到的内容为“COLLECTION REPORT”，将之后的内容导出表
                            System.out.println("输出第" + num + "张表:" + rowData);
                            XSSFSheet xssfSheet = xssfWorkbook.createSheet("Sheet"+ num);
                            CellRangeAddress callRangeAddress = new CellRangeAddress(0,0,2,3);
                            xssfSheet.addMergedRegion(callRangeAddress);
                            xssfSheet.setColumnWidth(0, 6000);
                            xssfSheet.setColumnWidth(1, 10000);
                            xssfSheet.setColumnWidth(2, 5000);
                            xssfSheet.setColumnWidth(3, 5000);
                            xssfSheet.setColumnWidth(4, 5000);
                            xssfSheet.setColumnWidth(5, 5000);
                            xssfSheet.setColumnWidth(6, 5000);
                            num ++;

                            int rowNum = 1;
                            for (int j = r + 1; j <= sheet_ETL.getLastRowNum(); j++) {
                                Row rows = sheet_ETL.getRow(j);
                                XSSFRow row0 = xssfSheet.createRow(0);//从第0行开始导出数据
                                row0.createCell(2).setCellValue(rowData);//导出第0行
                                if (rows != null){
                                    String data = "";
                                    if (rows.getCell(0) != null){
                                        rows.getCell(0).setCellType(Cell.CELL_TYPE_STRING);
                                        data = rows.getCell(0).getStringCellValue();
                                    }
                                    if (!"COLLECTION REPORT".equals(data)){
                                        XSSFRow row1 = xssfSheet.createRow(rowNum);
                                        if (rows.getCell(0) != null){
                                            rows.getCell(0).setCellType(Cell.CELL_TYPE_STRING);
                                            row1.createCell(0).setCellType(Cell.CELL_TYPE_STRING);
                                            row1.createCell(0).setCellValue(rows.getCell(0).getStringCellValue());
                                        }else {
                                            row1.createCell(0).setCellType(Cell.CELL_TYPE_STRING);
                                            row1.createCell(0).setCellValue("");
                                        }
                                        if (rows.getCell(1) != null){
                                            rows.getCell(1).setCellType(Cell.CELL_TYPE_STRING);
                                            row1.createCell(1).setCellType(Cell.CELL_TYPE_STRING);
                                            row1.createCell(1).setCellValue(rows.getCell(1).getStringCellValue());
                                        } else {
                                            row1.createCell(1).setCellType(Cell.CELL_TYPE_STRING);
                                            row1.createCell(1).setCellValue("");
                                        }
                                        if (rows.getCell(2) != null){
                                            if (rows.getCell(2).getCellType() == 0){
                                                rows.getCell(2).setCellType(Cell.CELL_TYPE_STRING);
                                                row1.createCell(2).setCellType(Cell.CELL_TYPE_STRING);
                                                String str = rows.getCell(2).getStringCellValue();
                                                if (str.contains(".")){
                                                    row1.createCell(2).setCellValue(convertData(str));
                                                }else{
                                                    row1.createCell(2).setCellValue(str);
                                                }
                                            }else {
                                                rows.getCell(2).setCellType(Cell.CELL_TYPE_STRING);
                                                row1.createCell(2).setCellType(Cell.CELL_TYPE_STRING);
                                                String str = rows.getCell(2).getStringCellValue();
                                                Pattern p = Pattern.compile("[a-zA-Z]");
                                                if (str.contains(".") && !p.matcher(str).find()){
                                                    row1.createCell(2).setCellValue(convertData(str));
                                                }else {
                                                    row1.createCell(2).setCellValue(str);
                                                }
                                            }
                                        }else {
                                            row1.createCell(2).setCellType(Cell.CELL_TYPE_STRING);
                                            row1.createCell(2).setCellValue("");
                                        }
                                        if (rows.getCell(3) != null){
                                            if (rows.getCell(3).getCellType() == 0){
                                                if (DateUtil.isCellDateFormatted(rows.getCell(3))){//判断是否是时间类型
                                                    String date = sdf.format(rows.getCell(3).getDateCellValue());
                                                    row1.createCell(3).setCellValue(date);
                                                }else {
                                                    rows.getCell(3).setCellType(Cell.CELL_TYPE_STRING);
                                                    row1.createCell(3).setCellType(Cell.CELL_TYPE_STRING);
                                                    String str = rows.getCell(3).getStringCellValue();
                                                    if (str.contains(".")){
                                                        row1.createCell(3).setCellValue(convertData(str));
                                                    }else{
                                                        row1.createCell(3).setCellValue(str);
                                                    }
                                                }
                                            }else {
                                                rows.getCell(3).setCellType(Cell.CELL_TYPE_STRING);
                                                row1.createCell(3).setCellType(Cell.CELL_TYPE_STRING);
                                                String str = rows.getCell(3).getStringCellValue();
                                                Pattern p = Pattern.compile("[a-zA-Z]");
                                                if (str.contains(".") && !p.matcher(str).find()){
                                                    row1.createCell(3).setCellValue(convertData(str));
                                                }else {
                                                    row1.createCell(3).setCellValue(str);
                                                }
                                            }
                                        }else {
                                            row1.createCell(3).setCellType(Cell.CELL_TYPE_STRING);
                                            row1.createCell(3).setCellValue("");
                                        }
                                        if (rows.getCell(4) != null){
                                            if (rows.getCell(4).getCellType() == 0){
                                                rows.getCell(4).setCellType(Cell.CELL_TYPE_STRING);
                                                row1.createCell(4).setCellType(Cell.CELL_TYPE_STRING);
                                                String str = rows.getCell(4).getStringCellValue();
                                                if (str.contains(".")){
                                                    row1.createCell(4).setCellValue(convertData(str));
                                                }else{
                                                    row1.createCell(4).setCellValue(str);
                                                }
                                            }else {
                                                rows.getCell(4).setCellType(Cell.CELL_TYPE_STRING);
                                                row1.createCell(4).setCellType(Cell.CELL_TYPE_STRING);
                                                String str = rows.getCell(4).getStringCellValue();
                                                Pattern p = Pattern.compile("[a-zA-Z]");
                                                if (str.contains(".") && !p.matcher(str).find()){
                                                    row1.createCell(4).setCellValue(convertData(str));
                                                }else {
                                                    row1.createCell(4).setCellValue(str);
                                                }
                                            }
                                        }else {
                                            row1.createCell(4).setCellType(Cell.CELL_TYPE_STRING);
                                            row1.createCell(4).setCellValue("");
                                        }
                                        if (rows.getCell(5) != null){
                                            rows.getCell(5).setCellType(Cell.CELL_TYPE_STRING);
                                            row1.createCell(5).setCellType(Cell.CELL_TYPE_STRING);
                                            row1.createCell(5).setCellValue(rows.getCell(5).getStringCellValue());
                                        }else {
                                            row1.createCell(5).setCellType(Cell.CELL_TYPE_STRING);
                                            row1.createCell(5).setCellValue("");
                                        }
                                        if (rows.getCell(6) != null){
                                            if (rows.getCell(6).getCellType() == 0){
                                                if (DateUtil.isCellDateFormatted(rows.getCell(6))){//判断是否是时间类型
                                                    String date = sdf.format(rows.getCell(6).getDateCellValue());
                                                    row1.createCell(6).setCellValue(date);
                                                }
                                            }else {
                                                rows.getCell(6).setCellType(Cell.CELL_TYPE_STRING);
                                                row1.createCell(6).setCellType(Cell.CELL_TYPE_STRING);
                                                row1.createCell(6).setCellValue(rows.getCell(6).getStringCellValue());
                                            }
                                        }else {
                                            row1.createCell(6).setCellType(Cell.CELL_TYPE_STRING);
                                            row1.createCell(6).setCellValue("");
                                        }
                                        rowNum++;
                                    }else {
                                        index = j;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return xssfWorkbook;
    }
}
