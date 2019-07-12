package HttpUtil;

import DataClean.TxtData;
import org.apache.commons.collections.bag.SynchronizedSortedBag;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 德国AR核销数据处理
 * 读取TXT文件数据内容，提取有效数据，输出表格
 */
public class Germany_AR {
    public static void main(String[] args) {
//        String filePath = args[0];//TXT文档路径
//        String Excel_filePath = args[1];//生成表路径
        String filePath = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\德国AR\\test\\0000815530_H20190330 (1).txt";
        String Excel_filePath = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AR\\德国AR\\test\\Excel.xlsx";//生成表路径
        //获取TXT文档内容,提取有效数据
        List<String> dataList = ReadTxtContent(filePath);
        //有效数据转换为实体数据
        List<TxtData> txtDataList = HandleData(dataList);
        //生成路径
        System.out.println("输出：" + Excel_filePath);
        //生成数据表
        excelOutput_Data(txtDataList,Excel_filePath);
        System.out.println("运行结束！");
    }

    /**
     * 导出Excel表格
     * @param txtDataList
     * @param filePath
     */
    public static void excelOutput_Data(List<TxtData> txtDataList, String filePath){
        //创建表格
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        //定义第一个sheet页
        XSSFSheet xssfSheet = xssfWorkbook.createSheet("Sheet1");
        //第一个sheet页数据
        Row row0 = xssfSheet.createRow(0);
        String[] headers = new String[]{"Beleg","Reference","Belegdatum","SAP Amount", "Discount", "Amount Paid","Summe"};
        for (int i = 0; i < headers.length; i++) {
            XSSFCell cell = (XSSFCell) row0.createCell(i);
            XSSFRichTextString text = new XSSFRichTextString(headers[i]);
            cell.setCellValue(text);
            xssfSheet.setColumnWidth(i, 5000);
        }
        //第一个Sheet页导出数据
        int rowNum = 1;
        for (TxtData txtData : txtDataList){
            XSSFRow row = xssfSheet.createRow(rowNum);
            xssfSheet.setColumnWidth(rowNum, 5000);
            row.createCell(0).setCellValue(txtData.getBeleg());
            row.createCell(1).setCellValue(txtData.getReferenz());
            row.createCell(2).setCellValue(txtData.getBelegDatum());
            row.createCell(3).setCellValue(txtData.getBruttoBetrag());
            row.createCell(4).setCellValue(txtData.getSkonto());
            row.createCell(5).setCellValue(txtData.getZahlBetrag());
            row.createCell(6).setCellValue(txtData.getSum());
            rowNum++;
        }
        try {
            FileOutputStream fileOutputStream = new FileOutputStream(filePath);
            xssfWorkbook.write(fileOutputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 将有效数据转换为实体类型
     * 由于TXT中获取到的数字类型的数据值的分位符以及小数点显示有误
     * 如：123.233,00
     * 需要将分位符以及小数点矫正
     * @param dataList
     * @return
     */
    public static List HandleData(List<String> dataList){
        List<TxtData> txtDataList = new ArrayList<>();
        for (String data : dataList){
            TxtData txtData = new TxtData();
            //将字符串中的“.”全部换为“,”
            data = data.replace(".",",");
            //拆分数据，调用方法矫正分位符以及小数点
            String[] vls = data.split(" ");
            List<String> list = Arrays.asList(vls);
//            System.out.println("在清理之前的Beleg是:"+list.get(0)+"在清理之前的Referenz是:"+list.get(1)+"在清理之前的BelegDatum是:"+list.get(2)+"在清理之前的BelegDatum是:"+list.get(3)+"在清理之前的BruttoBetrag是:"+list.get(4)+"在清理之前的ZahlBetrag是:"+list.get(5));

            txtData.setBeleg(list.get(0));
            String Referenz = list.get(1);
            //如果是RT30，则切割补零
            if ("RT30".equals(list.get(6))){

                if (Referenz.contains("/")){
                    Referenz=Referenz.substring(5);
                    int length = Referenz.length();
                    int ZeroNum = 10-length;
                    StringBuilder stringBuilder = new StringBuilder("");
                    for (int i =1;i<=ZeroNum;i++){
                        stringBuilder=stringBuilder.append("0");
                    }
                    String str = stringBuilder.toString();
                    Referenz = Referenz.substring(0,2)+str+Referenz.substring(2);
                    System.out.println(Referenz);
                }else {
                    Referenz=Referenz.substring(4);
                    int length = Referenz.length();
                    int ZeroNum = 10-length;
                    StringBuilder stringBuilder = new StringBuilder("");
                    for (int i =1;i<=ZeroNum;i++){
                        stringBuilder=stringBuilder.append("0");
                    }
                    String str = stringBuilder.toString();
                    Referenz = Referenz.substring(0,2)+str+Referenz.substring(2);
                    System.out.println(Referenz);
                }
            }
            txtData.setReferenz(Referenz);
            txtData.setBelegDatum(list.get(2));
            txtData.setBruttoBetrag(clearData(list.get(3)));
            txtData.setSkonto(clearData(list.get(4)));
            txtData.setZahlBetrag(clearData(list.get(5)));
            txtData.setSum(list.get(6));
            txtDataList.add(txtData);
        }
        return txtDataList;
    }

    /**
     * 矫正数据的分位符，以及小数点
     * @param str
     * @return
     */
    public static String clearData(String str){
        if (!"".equals(str)){
            //判断数据中是否有“-”，若有，则将“-”放到数据的开头位置
            if (str.contains("-")){
                str = "-" + str.replace("-","");
            }
            StringBuilder sb = new StringBuilder(str);
            int i = str.lastIndexOf(",");
            sb.replace(i,i+1 ,".");
            str = sb.toString();
        }
        return str;
    }

    /**
     * 获取txt文档的内容，提取有效数据以及编号
     * @param filePath
     * @return
     */
    public static List  ReadTxtContent(String filePath){
        File file = new File(filePath);
        List<String> dataList = new ArrayList<>();
        try{
            BufferedReader br = new BufferedReader(new FileReader(file));//构造一个BufferedReader类来读取文件
            String s = null;
            String  number = "";
            while((s = br.readLine())!= null){//使用readLine方法，一次读一行
                String str = System.lineSeparator() + s;//每一行的内容
                //将连续的多个空格换为一个空格
                Pattern p = Pattern.compile("\\s+");
                Matcher m = p.matcher(str.trim());
                str = m.replaceAll(" ");

                //#################################
                //将每一行字符串的空格、空字符去掉
                String strs = str.replace(" ","").trim();
                //匹配满足条件的编号（例：-M016）
                Pattern pp = Pattern.compile("(-[A-Z]{1}+[0-9]{3})|(-[A-Z]{2}+[0-9]{2})");
                Matcher mm = pp.matcher(strs);
                //获取编号
                String str2 = "";
                while(mm.find()){
                    str2 = str2 + mm.group();
                }

                //判断是否获取到有效编号，有则将获取到的编号去“-”
                if (!"".equals(str2)){
                    number = str2.replace("-","");
                }
                //将编号拼接到字符串后面
                str = str + " " + number;
                //###################################

                //统计空格的个数，过滤掉部分数据
                int count = 0;
                for (int i =0; i < str.length(); i++){
                    char tem = str.charAt(i);
                    if (tem == ' '){// 空格
                        count++;
                    }
                }
                //同时，将包含有“,”的字符串数据取出（所需有效数据都包含逗号）
                if (count >= 6 && str.contains(",")){
                    Pattern pattern = Pattern.compile("[0-9]*");
                    String index = String.valueOf(str.charAt(0));
                    if (pattern.matcher(index).matches()){
//                        System.out.println("往List里面放的字符串是:"+str);
                        dataList.add(str);
                    }
                }
            }
            br.close();
        }catch(Exception e){
            e.printStackTrace();
        }
        return dataList;
    }
}
