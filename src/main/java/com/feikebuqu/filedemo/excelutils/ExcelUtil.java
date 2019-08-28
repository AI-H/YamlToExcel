package com.feikebuqu.filedemo.excelutils;


import com.feikebuqu.filedemo.bean.ExcelBean;
import com.feikebuqu.filedemo.bean.ExcelField;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by HuangLiTe on 2017/10/31.
 */
public class ExcelUtil {
    public  static <E> Workbook createTsWorkBook(List<E> list, Class<E> clazz) throws Exception{
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("sheet1");
        Field[] fieldArray = clazz.getDeclaredFields();
        List<String> getParamList = new ArrayList<>();
        List<String> columnNameList = new ArrayList<>();
        for (int i = 0; i < fieldArray.length; i++) {
            sheet.setColumnWidth(i, (50 * 150));//设置列宽度
            Field field = fieldArray[i];
            String name = field.getName();
            ExcelField myFieldAnnotation = field.getAnnotation(ExcelField.class);
            if (myFieldAnnotation != null) {
                columnNameList.add(myFieldAnnotation.value());
                getParamList.add("get" + name.substring(0, 1).toUpperCase() + name.substring(1));
            }
        }
        Row row = sheet.createRow(0);
        CellStyle csTitle = wb.createCellStyle();
        CellStyle csContent = wb.createCellStyle();
        Font fTitle = wb.createFont();
        Font fContent = wb.createFont();
        fTitle.setFontHeightInPoints((short) 10);
        fTitle.setColor(IndexedColors.BLACK.getIndex());
        fTitle.setBoldweight(Font.BOLDWEIGHT_BOLD);
        fContent.setFontHeightInPoints((short) 10);
        fContent.setColor(IndexedColors.BLACK.getIndex());
        csTitle.setFont(fTitle);
        csTitle.setBorderLeft(CellStyle.BORDER_THIN);
        csTitle.setBorderRight(CellStyle.BORDER_THIN);
        csTitle.setBorderTop(CellStyle.BORDER_THIN);
        csTitle.setBorderBottom(CellStyle.BORDER_THIN);
        csTitle.setAlignment(CellStyle.ALIGN_CENTER);

        csContent.setFont(fContent);
        csContent.setBorderLeft(CellStyle.BORDER_THIN);
        csContent.setBorderRight(CellStyle.BORDER_THIN);
        csContent.setBorderTop(CellStyle.BORDER_THIN);
        csContent.setBorderBottom(CellStyle.BORDER_THIN);
        csContent.setAlignment(CellStyle.ALIGN_CENTER);

        for (int i = 0; i < columnNameList.size(); i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(columnNameList.get(i));
            cell.setCellStyle(csTitle);
        }

        for (int i = 0; i < list.size(); i++) {
            Row rowContent = sheet.createRow(i + 1);
            E bean = list.get(i);
            for (int j = 0; j < getParamList.size(); j++) {
                Cell cell = rowContent.createCell(j);
                Object value = ReflectionUtil.invokeMethod(bean, getParamList.get(j));
                cell.setCellValue(value == null ? "" : value.toString());
                cell.setCellStyle(csContent);
            }
        }
        return wb;
    }

    public static <E> Workbook createTsWorkBookList(List<ExcelBean> excelBeanList) throws Exception{
        Workbook wb = new HSSFWorkbook();
        for (ExcelBean excelBean : excelBeanList) {
            Sheet sheet = wb.createSheet(excelBean.getTableName());
            Field[] fieldArray = excelBean.getClazz().getDeclaredFields();
            List<String> getParamList = new ArrayList<>();
            List<String> columnNameList = new ArrayList<>();
            for (int i = 0; i < fieldArray.length; i++) {
                sheet.setColumnWidth(i, (50 * 150));//设置列宽度
                Field field = fieldArray[i];
                String name = field.getName();
                ExcelField myFieldAnnotation = field.getAnnotation(ExcelField.class);
                if (myFieldAnnotation != null) {
                    columnNameList.add(myFieldAnnotation.value());
                    getParamList.add("get" + name.substring(0, 1).toUpperCase() + name.substring(1));
                }
            }
            Row row = sheet.createRow(0);
            CellStyle csTitle = wb.createCellStyle();
            CellStyle csContent = wb.createCellStyle();
            Font fTitle = wb.createFont();
            Font fContent = wb.createFont();
            fTitle.setFontHeightInPoints((short) 10);
            fTitle.setColor(IndexedColors.BLACK.getIndex());
            fTitle.setBoldweight(Font.BOLDWEIGHT_BOLD);
            fContent.setFontHeightInPoints((short) 10);
            fContent.setColor(IndexedColors.BLACK.getIndex());
            csTitle.setFont(fTitle);
            csTitle.setBorderLeft(CellStyle.BORDER_THIN);
            csTitle.setBorderRight(CellStyle.BORDER_THIN);
            csTitle.setBorderTop(CellStyle.BORDER_THIN);
            csTitle.setBorderBottom(CellStyle.BORDER_THIN);
            csTitle.setAlignment(CellStyle.ALIGN_CENTER);

            csContent.setFont(fContent);
            csContent.setBorderLeft(CellStyle.BORDER_THIN);
            csContent.setBorderRight(CellStyle.BORDER_THIN);
            csContent.setBorderTop(CellStyle.BORDER_THIN);
            csContent.setBorderBottom(CellStyle.BORDER_THIN);
            csContent.setAlignment(CellStyle.ALIGN_CENTER);

            for (int i = 0; i < columnNameList.size(); i++) {
                Cell cell = row.createCell(i);
                cell.setCellValue(columnNameList.get(i));
                cell.setCellStyle(csTitle);
            }

            List<E> list = excelBean.getList();
            for (int i = 0; i < list.size(); i++) {
                Row rowContent = sheet.createRow(i + 1);
                E bean = list.get(i);
                for (int j = 0; j < getParamList.size(); j++) {
                    Cell cell = rowContent.createCell(j);
                    Object value = ReflectionUtil.invokeMethod(bean, getParamList.get(j));
                    cell.setCellValue(value == null ? "" : value.toString());
                    cell.setCellStyle(csContent);
                }
            }
            if (excelBean.getLastRowName() != null) {
                Row rowContent = sheet.createRow(list.size() + 1);
                Cell cell = rowContent.createCell(0);
                cell.setCellValue(excelBean.getLastRowName());
                cell.setCellStyle(csContent);
                cell = rowContent.createCell(1);
                cell.setCellValue(excelBean.getLastRowValue());
                cell.setCellStyle(csContent);
            }
        }
        return wb;
    }

    /**
     * 导出excel
     * @param response
     * @param list
     * @param clazz
     * @param fileName
     * @param <E>
     * @throws Exception
     */
    public static <E> void exportExcel(HttpServletResponse response, List<E> list, Class<E> clazz, String fileName)
            throws Exception {
        Workbook wb = createTsWorkBook(list, clazz);
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        wb.write(os);
        byte[] content = os.toByteArray();
        InputStream is = new ByteArrayInputStream(content);
        // 设置response参数，可以打开下载页面
        response.reset();
        response.setContentType("application/vnd.ms-excel;charset=utf-8");
        response.setHeader("Content-Disposition", "attachment;filename="
                + new String((fileName + ".xls").getBytes(), "iso-8859-1"));
        ServletOutputStream out = response.getOutputStream();
        BufferedInputStream bis = null;
        BufferedOutputStream bos = null;
        try {
            bis = new BufferedInputStream(is);
            bos = new BufferedOutputStream(out);
            byte[] buff = new byte[2048];
            int bytesRead;
            while (-1 != (bytesRead = bis.read(buff, 0, buff.length))) {
                bos.write(buff, 0, bytesRead);
            }
        } catch (final IOException e) {
            throw e;
        } finally {
            if (bis != null)
                bis.close();
            if (bos != null)
                bos.close();
        }
    }

    /**
     * 导出excel - 多列表
     * @param response
     * @param excelBeanList
     * @param fileName
     * @throws Exception
     */
    public static void exportExcelList(HttpServletResponse response, List<ExcelBean> excelBeanList, String fileName)
            throws Exception {
        Workbook wb = createTsWorkBookList(excelBeanList);
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        wb.write(os);
        byte[] content = os.toByteArray();
        InputStream is = new ByteArrayInputStream(content);
        // 设置response参数，可以打开下载页面
        response.reset();
        response.setContentType("application/vnd.ms-excel;charset=utf-8");
        response.setHeader("Content-Disposition", "attachment;filename="
                + new String((fileName + ".xls").getBytes(), "iso-8859-1"));
        ServletOutputStream out = response.getOutputStream();
        BufferedInputStream bis = null;
        BufferedOutputStream bos = null;
        try {
            bis = new BufferedInputStream(is);
            bos = new BufferedOutputStream(out);
            byte[] buff = new byte[2048];
            int bytesRead;
            while (-1 != (bytesRead = bis.read(buff, 0, buff.length))) {
                bos.write(buff, 0, bytesRead);
            }
        } catch (final IOException e) {
            throw e;
        } finally {
            if (bis != null)
                bis.close();
            if (bos != null)
                bos.close();
        }
    }

    /**
     * 读取excel文件内容
     * @param inputStream FileInputStream
     * @param fileType
     * @param <E>
     * @return
     * @throws Exception
     */
    public static <E> List<E> readExcel(InputStream inputStream, String fileType, Class<E> clazz) throws Exception {
        Workbook wbdq = null;
        if (".xls".equals(fileType)) {
            wbdq = new HSSFWorkbook(inputStream);
        } else {
            wbdq = new XSSFWorkbook(inputStream);
        }
        Sheet sheetdq = wbdq.getSheetAt(0);
        Field[] fieldArray = clazz.getDeclaredFields();
        List<String> setParamList = new ArrayList<>();
        for (int i = 0; i < fieldArray.length; i++) {
            Field field = fieldArray[i];
            String name = field.getName();
            setParamList.add("set" + name.substring(0, 1).toUpperCase() + name.substring(1));
        }
        List<E> list = new ArrayList<E>();
        for (int i = 1; i < sheetdq.getLastRowNum() + 1; i++) {
            Row rowdq = sheetdq.getRow(i);
            if (rowdq == null){
                break;
            }
            E bean = clazz.newInstance();
            for (int j = 0; j < setParamList.size(); j++) {
                Cell cell = rowdq.getCell(j);
                String value;
                if (cell != null && cell.getCellType() == Cell.CELL_TYPE_NUMERIC){
                    value = new DecimalFormat("0.00").format(cell.getNumericCellValue()).trim();
                } else if (cell != null && cell.getCellType() == Cell.CELL_TYPE_STRING){
                    value = cell.getStringCellValue().trim();
                } else {
                    System.err.println("批量导入处理到:" + i + "行为空或单元格类型错误");
                    continue;
                }
                ReflectionUtil.invokeMethod(bean, setParamList.get(j), value);
            }
            list.add(bean);
        }

        return list;
    }

    /**
     * 读取excel文件内容 单列
     * @param inputStream FileInputStream
     * @param fileType
     * @return
     * @throws Exception
     */
    public static List<String> readExcel(InputStream inputStream, String fileType) throws Exception {
        Workbook wbdq = null;
        if (".xls".equals(fileType)) {
            wbdq = new HSSFWorkbook(inputStream);
        } else {
            wbdq = new XSSFWorkbook(inputStream);
        }
        Sheet sheetdq = wbdq.getSheetAt(0);
        List<String> list = new ArrayList<String>();
        for (int i = 0; i < sheetdq.getLastRowNum() + 1; i++) {
            Row rowdq = sheetdq.getRow(i);
            if (rowdq == null){
                break;
            }
            Cell cell = rowdq.getCell(0);
            String value;
            if (cell != null && cell.getCellType() == Cell.CELL_TYPE_NUMERIC){
                value = new DecimalFormat("0.00").format(cell.getNumericCellValue()).trim();
            } else if (cell != null && cell.getCellType() == Cell.CELL_TYPE_STRING){
                value = cell.getStringCellValue().trim();
            } else {
                System.err.println("批量导入处理到:" + i + "行为空或单元格类型错误");
                continue;
            }
            list.add(value);
        }

         return list;
    }

    public static boolean checkFile(String fileName){
        boolean flag=false;
        String suffixList="xls,xlsx";
        String suffix=fileName.substring(fileName.lastIndexOf(".")+1, fileName.length());
        if(suffixList.contains(suffix.trim().toLowerCase())){
            flag=true;
        }
        return flag;
    }

}
