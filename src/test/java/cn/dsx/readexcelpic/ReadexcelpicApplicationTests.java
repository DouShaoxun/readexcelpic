package cn.dsx.readexcelpic;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;
import org.apache.poi.ss.usermodel.PictureData;

import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@SpringBootTest
class ReadexcelpicApplicationTests {

    @Test
    void contextLoads() throws Exception {
        File directory = new File("src/main/resources");
        String courseFile = directory.getCanonicalPath();
        String excelPath = courseFile + "/excel/test.xls";

        InputStream inp = new FileInputStream(excelPath);
        HSSFWorkbook workbook = (HSSFWorkbook) WorkbookFactory.create(inp);
        List<HSSFPictureData> pictures = workbook.getAllPictures();
        HSSFSheet sheet = (HSSFSheet) workbook.getSheetAt(0);

        Map<Integer, HSSFPictureData> picMap = getPicMap(workbook);
        int i = 0;
        for (HSSFShape shape : sheet.getDrawingPatriarch().getChildren()) {
            HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();

            if (shape instanceof HSSFPicture) {
                HSSFPicture pic = (HSSFPicture) shape;
                int pictureIndex = pic.getPictureIndex() - 1;
                HSSFPictureData picData = pictures.get(pictureIndex);
                savePic(anchor.getRow1(), anchor.getCol1(), picData, courseFile + "/excel/");
            }
            i++;
        }
        System.out.println(picMap.size());
    }

    /**
     * 写入磁盘
     * @param row
     * @param col
     * @param pic
     * @param filePath
     * @throws Exception
     */
    private void savePic(int row, int col, PictureData pic, String filePath) throws Exception {
        System.out.println(row + ":" + col);
        String ext = pic.suggestFileExtension();
        // byte 数据
        byte[] data = pic.getData();
        if (ext.equals("jpeg")) {
            FileOutputStream out = new FileOutputStream(filePath + row + "_" + col + ".jpg");
            out.write(data);
            out.close();
        }
        if (ext.equals("png")) {
            FileOutputStream out = new FileOutputStream(filePath + row + "_" + col + ".png");
            out.write(data);
            out.close();
        }
    }


    /**
     * 获取excel 所有图片
     *
     * @param workbook
     * @return
     */
    private Map<Integer, HSSFPictureData> getPicMap(HSSFWorkbook workbook) {
        Map<Integer, HSSFPictureData> picMap = new HashMap<Integer, HSSFPictureData>();

        List<HSSFPictureData> pictures = workbook.getAllPictures();
        HSSFSheet sheet = (HSSFSheet) workbook.getSheetAt(0);

        for (HSSFShape shape : sheet.getDrawingPatriarch().getChildren()) {
            HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
            if (shape instanceof HSSFPicture) {
                HSSFPicture pic = (HSSFPicture) shape;
                HSSFPictureData picData = pictures.get(pic.getPictureIndex() - 1);

                picMap.put(anchor.getRow1(), picData);
            } else {
                picMap.put(anchor.getRow1(), null);  //非图片数据则插入null
            }
        }

        return picMap;
    }
}
