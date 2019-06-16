package com.zhzh;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.common.IOUtil;
import org.junit.Test;

import java.io.*;

/**
 * ${Author}: jason.zhao
 * 2019/6/10 20:49
 **/
public class POITest01 {

    @Test
    public void test01() throws Exception {
        /**
         * HSSF提供读写Microsoft Excel XLS格式档案的功能。
         XSSF提供读写Microsoft Excel OOXML XLSX格式档案的功能。
         HWPF提供读写Microsoft Word DOC格式档案的功能。
         HSLF提供读写Microsoft PowerPoint格式档案的功能。
         HDGF提供读Microsoft Visio格式档案的功能。
         HPBF提供读Microsoft Publisher格式档案的功能。
         HSMF提供读Microsoft Outlook格式档案的功能。
         */
        //创建工作簿
        Workbook wb = new XSSFWorkbook();
        //创建表单
        Sheet sheet = wb.createSheet("test");
        //3.创建行对象，从0开始
        Row row = sheet.createRow(3);
        //4.创建单元格，从0开始
        Cell cell = row.createCell(0);
        //写入内容
        cell.setCellValue("zhzh");

        CellStyle cellStyle = wb.createCellStyle();

        //设置边框
        cellStyle.setBorderBottom(BorderStyle.DASH_DOT);//下边框
        cellStyle.setBorderTop(BorderStyle.HAIR);//上边框

        Font font = wb.createFont();
        font.setFontName("华文行楷");
        //设置字号
        font.setFontHeightInPoints((short)28);
        cellStyle.setFont(font);
        //设置宽高
        sheet.setColumnWidth(0, 31 * 256);//设置第一列的宽度是31个字符宽度
        row.setHeightInPoints(50);//设置行的高度是50个点

        cellStyle.setAlignment(HorizontalAlignment.CENTER);//水平居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
        cell.setCellStyle(cellStyle);

        CellRangeAddress region =new CellRangeAddress(0, 3, 0, 2);
        sheet.addMergedRegion(region);

        FileOutputStream fos = new FileOutputStream("E:\\test.xlsx");
        ByteArrayOutputStream baos = new ByteArrayOutputStream();

        //写入文件
        wb.write(baos);
        baos.writeTo(fos);
        byte[] bytes = baos.toByteArray();
        fos.close();

    }

    /**
     * 绘制图形
     */
    @Test
    public void  test02() throws IOException {
        //1.创建workbook工作簿
        Workbook wb = new XSSFWorkbook();
        //2.创建表单Sheet
        Sheet sheet = wb.createSheet("test");
        FileInputStream fis =
                new FileInputStream("C:\\Users\\Administrator\\Desktop\\1.png");
        byte[] bytes = IOUtils.toByteArray(fis);
        fis.read(bytes);
        //向Excel添加一张图片,并返回该图片在Excel中的图片集合中的下标
        int pictureIdx = wb.addPicture(bytes,Workbook.PICTURE_TYPE_JPEG);
        //绘图工具类
        CreationHelper helper = wb.getCreationHelper();
        //创建一个绘图对象
        Drawing<?> drawingPatriarch = sheet.createDrawingPatriarch();
        //创建锚点,设置图片坐标
        ClientAnchor anchor = helper.createClientAnchor();
        anchor.setCol1(0);
        anchor.setRow1(0);

        //创建图片
        Picture picture = drawingPatriarch.createPicture(anchor, pictureIdx);
        picture.resize();

        //6.文件流
        FileOutputStream fos = new FileOutputStream("E:\\test.xlsx");
        //7.写入文件
        wb.write(fos);
        fos.close();

    }
    @Test
    public void test03() throws IOException {
        //1.创建workbook工作簿

        Workbook wb = new XSSFWorkbook("D:\\project\\test-poi\\src\\main\\resources\\excel\\demo.xlsx");
        //2.获取sheet 从0开始
        Sheet sheet = wb.getSheetAt(0);

        int totalRowNum = sheet.getLastRowNum();
        Row row = null;
        Cell cell = null;

        for(int rowNum=0;rowNum<=sheet.getLastRowNum();rowNum++){
            row = sheet.getRow(rowNum);
            StringBuilder sb = new StringBuilder();
            for(int cellNum=2;cellNum<row.getLastCellNum();cellNum++){
                cell = row.getCell(cellNum);
               sb.append(getValue(cell)).append("-");

            }
            System.out.println(sb.toString());
        }
    }
    //获取数据
    private static Object getValue(Cell cell) {
        Object value = null;
        switch (cell.getCellType()) {
            case STRING: //字符串类型
                value = cell.getStringCellValue();
                break;
            case BOOLEAN: //boolean类型
                value = cell.getBooleanCellValue();
                break;
            case NUMERIC: //数字类型（包含日期和普通数字）
                if (DateUtil.isCellDateFormatted(cell)) {
                    value = cell.getDateCellValue();
                } else {
                    value = cell.getNumericCellValue();
                }
                break;
            case FORMULA: //公式类型
                value = cell.getCellFormula();
                break;
            default:
                break;
        }
        return value;
    }

}
