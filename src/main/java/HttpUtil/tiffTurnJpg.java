package HttpUtil;


import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import javax.imageio.ImageIO;
import javax.media.jai.JAI;
import javax.media.jai.RenderedOp;

import com.sun.image.codec.jpeg.JPEGHuffmanTable;
import com.sun.image.codec.jpeg.JPEGQTable;
import com.sun.media.jai.codec.BMPEncodeParam;
import com.sun.media.jai.codec.ImageCodec;
import com.sun.media.jai.codec.ImageEncodeParam;
import com.sun.media.jai.codec.ImageEncoder;
import com.sun.media.jai.codec.JPEGEncodeParam;
//import com.sun.image.codec.jpeg.JPEGEncodeParam;
public class tiffTurnJpg {
    public static void main(String[] args){
//        String XDensity = args[0];
//        String YDensity = args[1];
//        String filePath=args[2];
//        String jpgOriginalPath = args[3];
        String XDensity = "3308";
        String YDensity= "4680";
        String filePath = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AP\\新加坡用horis的国家\\巴基斯坦\\Main\\3778853 INV PL.tif";
        String jpgOriginalPath="C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AP\\新加坡用horis的国家\\巴基斯坦\\Main\\3.jpg";
        String jpgPath = jpgOriginalPath.substring(0,jpgOriginalPath.lastIndexOf("."))+"S.jpg";
        RenderedOp file = JAI.create("fileload", filePath);//读取tiff图片文件
        OutputStream ops = null;
        try {
            ops = new FileOutputStream(jpgOriginalPath);
            //文件存储输出流
            JPEGEncodeParam param = new JPEGEncodeParam() {
            };
            ImageEncoder image = ImageCodec.createImageEncoder("JPEG", ops, (ImageEncodeParam) param); //指定输出格式
            //解析输出流进行输出
            image.encode(file);
            //关闭流
            ops.close();
            //转换成指定分辨率的jpg
            resizeImage(jpgOriginalPath,jpgOriginalPath,Integer.parseInt(XDensity),Integer.parseInt(YDensity));


        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("tiff转换jpg成功");
    }
    public static void resizeImage(String srcPath, String desPath,
                                   int width, int height) throws IOException {

        File srcFile = new File(srcPath);
        Image srcImg = ImageIO.read(srcFile);
        BufferedImage buffImg = null;
        buffImg = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        //使用TYPE_INT_RGB修改的图片会变色
        buffImg.getGraphics().drawImage(
                srcImg.getScaledInstance(width, height, Image.SCALE_SMOOTH), 0,
                0, null);

        ImageIO.write(buffImg, "JPEG", new File(desPath));
    }

}
