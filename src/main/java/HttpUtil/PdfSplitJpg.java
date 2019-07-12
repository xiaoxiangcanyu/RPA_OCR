package HttpUtil;

import com.sun.pdfview.PDFFile;
import com.sun.pdfview.PDFPage;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.ByteBuffer;
import java.nio.channels.FileChannel;
import java.util.ArrayList;
import java.util.List;

public class PdfSplitJpg {
    public String pdfTransferjpg(String file_pdf) {
//        String file_img = args[0];
        String file_img = file_pdf.substring(0, file_pdf.lastIndexOf("\\") + 1);
        String dstName = "";
//        String file_pdf = args[1];
//        String file_pdf = "C:\\Users\\songyu\\Desktop\\haier_rpa所有资料\\OCR_Data\\项目交接文档\\项目交接文档\\AP\\新加坡用horis的国家\\泰国工厂\\泰国工厂AP\\test\\1901265416.pdf";
        List<String> imgUrl = new ArrayList<>();
        try {
            String img_path = file_pdf.substring(file_pdf.lastIndexOf("\\") + 1, file_pdf.indexOf("."));
            File file = new File(file_pdf);
            RandomAccessFile raf = new RandomAccessFile(file, "r");
            FileChannel channel = raf.getChannel();
            ByteBuffer buf = channel.map(FileChannel.MapMode.READ_ONLY, 0, channel.size());
            PDFFile pdffile = new PDFFile(buf);

            String getPdfFilePath = file_img;
            //目录不存在，则创建目录
            File p = new File(getPdfFilePath);
            if (!p.exists()) {
                p.mkdir();
            }
            for (int i = 1; i <= pdffile.getNumPages(); i++) {
                PDFPage page = pdffile.getPage(i);
                int n = 6;
                Rectangle rect = new Rectangle(0, 0, (int) page.getBBox().getWidth(), (int) page.getBBox().getHeight());
                Image img = page.getImage(rect.width * n, rect.height * n, rect, null, true, true);

                BufferedImage tag = new BufferedImage(rect.width * n, rect.height * n, BufferedImage.TYPE_INT_RGB);
                tag.getGraphics().drawImage(img, 0, 0, rect.width * n, rect.height * n, null);
                //转换成功的图片路径dstName
                dstName = getPdfFilePath + File.separator + img_path + "_" + i + ".jpg";
                System.out.println(dstName);
                FileOutputStream out = new FileOutputStream(dstName); // 输出到文件流
                String formatName = dstName.substring(dstName.lastIndexOf(".") + 1);
                ImageIO.write(tag, /*"GIF"*/ formatName /* format desired */, new File(dstName) /* target */);
                imgUrl.add(dstName);
                out.close();
            }
            System.out.println("pdf转换图片成功！");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return dstName;
    }
//        return imgUrl;

}
