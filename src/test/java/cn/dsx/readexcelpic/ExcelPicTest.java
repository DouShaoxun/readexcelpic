package cn.dsx.readexcelpic;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.*;
import java.util.List;
import java.util.Map;

@SpringBootTest
public class ExcelPicTest {
    @Test
    void contextLoads() throws Exception {
        File directory = new File("src/main/resources");
        String courseFile = directory.getCanonicalPath();
        String excelPath = courseFile + "/excel/test.xlsx";
        //String excelPath = courseFile + "/excel/test.xls";

        InputStream inp = new FileInputStream(excelPath);
        Workbook workbook = WorkbookFactory.create(inp);
        //// 获取所有图片
        //List<? extends PictureData> allPictures = workbook.getAllPictures();
        //allPictures.forEach(pictureData -> {
        //    byte[] data = pictureData.getData();
        //});

        Sheet sheet = workbook.getSheetAt(0);
        Drawing drawingPatriarch = sheet.getDrawingPatriarch();

        // xlsx
        if (drawingPatriarch instanceof XSSFDrawing) {

            XSSFDrawing xssfDrawing = (XSSFDrawing) drawingPatriarch;
            List<XSSFShape> shapes = xssfDrawing.getShapes();
            shapes.forEach(shape -> {
                if (shape instanceof XSSFPicture) {
                    XSSFPicture picture = (XSSFPicture) shape;
                    XSSFPictureData pictureData = picture.getPictureData();
                    // 图片文件字节
                    byte[] data = pictureData.getData();
                    XSSFClientAnchor clientAnchor = ((XSSFPicture) shape).getClientAnchor();

                    int row1 = clientAnchor.getRow1();
                    int col1 = clientAnchor.getCol1();
                    savePic(clientAnchor.getRow1(), clientAnchor.getCol1(), pictureData, courseFile + "/excel/");
                }
            });
        }

        // xls
        if (drawingPatriarch instanceof HSSFPatriarch) {
            HSSFPatriarch patriarch = (HSSFPatriarch) drawingPatriarch;
            patriarch.getChildren().forEach(shape -> {
                HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
                if (shape instanceof HSSFPicture) {
                    HSSFPicture pic = (HSSFPicture) shape;
                    HSSFPictureData pictureData = pic.getPictureData();
                    savePic(anchor.getRow1(), anchor.getCol1(), pictureData, courseFile + "/excel/");
                }
            });

        }
    }


    /**
     * 写入磁盘
     *
     * @param row
     * @param col
     * @param pic
     * @param filePath
     * @throws Exception
     */
    private void savePic(int row, int col, PictureData pic, String filePath) {
        System.out.println(row + ":" + col);
        String ext = pic.suggestFileExtension();
        // byte 数据
        byte[] data = pic.getData();

        String suffix = ".jpg";
        if (ext.equals("jpeg")) {
            suffix = ".jpg";
        }
        if (ext.equals("png")) {
            suffix = ".png";
        }


        FileOutputStream out = null;
        try {
            out = new FileOutputStream(filePath + row + "_" + col + suffix);
            out.write(data);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (out != null) {
                    out.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }


    }

}
