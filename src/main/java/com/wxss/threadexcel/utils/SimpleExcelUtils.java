package com.wxss.threadexcel.utils;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;


public class SimpleExcelUtils {

    /**
     * excel下载或保存磁盘
     * @param workbook excel
     * @param fileName 文件名，为null时，默认取时间
     * @param response HTTP 的 HttpServletResponse
     * @param localOutputStream 磁盘输出流
     * @throws IOException
     */
    public static void write(Workbook workbook, String fileName, HttpServletResponse response,OutputStream localOutputStream) throws IOException {
        if (fileName == null){
            SimpleDateFormat df = new SimpleDateFormat("yyyyMMddHHmmss");//设置日期格式
            fileName = df.format(new Date());// new Date()为获取当前系统时间
        }
        // HTTP
        try {
            if (response != null){
                response.setHeader("Content-disposition", "attachment; filename=" + fileName + ".xlsx");
                response.setContentType("application/vnd.ms-excel;charset=UTF-8");
                workbook.write(response.getOutputStream());
            }else {
                // 磁盘
                workbook.write(localOutputStream);
            }
        } finally {
            if (localOutputStream != null){
                localOutputStream.close();
            }
            if (response != null && response.getOutputStream()!=null){
                response.getOutputStream().close();
            }

        }
    }

    public static Workbook createWorkbook(MultipartFile file)  {
        String filename = file.getOriginalFilename();
        Workbook workbook = null;
        try {
            if(filename.endsWith("xls")){
                //2003
                workbook = new HSSFWorkbook(file.getInputStream());
            }else if(filename.endsWith("xlsx")){
                //2007
                workbook = new XSSFWorkbook(file.getInputStream());
            }else{
                throw new Exception("文件不是Excel文件");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        return workbook;
    }

    /**
     * 获得Cell内容
     *
     * @param cell
     * @return
     */
    public static String getCellValue(Cell cell) {
        String value = "";
        if (cell != null) {
            // 以下是判断数据的类型
            switch (cell.getCellType()) {
                case NUMERIC: // 数字
                    value = cell.getNumericCellValue() + "";
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        Date date = cell.getDateCellValue();
                        if (date != null) {
                            value = new SimpleDateFormat("yyyy-MM-dd").format(date);
                        } else {
                            value = "";
                        }
                    } else {
                        value = new DecimalFormat("0").format(cell.getNumericCellValue());
                    }
                    break;
                case STRING: // 字符串
                    value = cell.getStringCellValue();
                    break;
                case BOOLEAN: // Boolean
                    value = cell.getBooleanCellValue() + "";
                    break;
                case FORMULA: // 公式
                    value = cell.getCellFormula() + "";
                    break;
                case BLANK: // 空值
                    value = "";
                    break;
                case ERROR: // 故障
                    value = "非法字符";
                    break;
                default:
                    value = "未知类型";
                    break;
            }
        }

        return value.trim();
    }
}
