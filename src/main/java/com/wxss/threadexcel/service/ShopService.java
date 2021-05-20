package com.wxss.threadexcel.service;

import com.wxss.threadexcel.common.excel.ParseResult;
import com.wxss.threadexcel.common.excel.Task;
import com.wxss.threadexcel.utils.SimpleExcelUtils;
import com.wxss.threadexcel.domain.entity.Shop;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.*;

@Service
@Slf4j
public class ShopService {
    private static final String TEMPLATE_FILE_DIR = "D:/Temp/Shop/";
    private static final String TEMPLATE_FLE_NAME = "ShopTemplate";
    private static final String TEMPLATE_FLE_SUFFIX = ".xlsx";
    List<Shop> shopList = new ArrayList<>();


    private static final String[] COLUMN_NAME = new String[]{"门店名称", "门店编号", "营业开始时间", "营业结束时间"};

    public void downloadTemplate(HttpServletResponse response) throws Exception {
        File templateFile = new File(TEMPLATE_FILE_DIR + TEMPLATE_FLE_NAME + TEMPLATE_FLE_SUFFIX);
        Workbook workbook = null;
        if (templateFile.exists()) {
            InputStream inputStream = new FileInputStream(templateFile);
            workbook = WorkbookFactory.create(inputStream);
        } else {
            // 模板文件不存在， 需要新创建模板文件
            // 第一步，创建一个webbook，对应一个Excel文件
            workbook = new XSSFWorkbook();
            // 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
            Sheet sheet = workbook.createSheet("Sheet1");
            // 第三步，设置标题列
            Row row = sheet.createRow(0);
            // 第四步，创建单元格，并设置值表头 设置表头居中
            CellStyle style = workbook.createCellStyle();
            style.setAlignment(HorizontalAlignment.CENTER); // 创建一个居中格式
            Cell cell = row.createCell(0);
            cell.setCellValue(COLUMN_NAME[0]);
            cell.setCellStyle(style);
            cell = row.createCell(1);
            cell.setCellValue(COLUMN_NAME[1]);
            cell.setCellStyle(style);
            cell = row.createCell(2);
            cell.setCellValue(COLUMN_NAME[2]);
            cell.setCellStyle(style);
            cell = row.createCell(3);
            cell.setCellValue(COLUMN_NAME[3]);
            cell.setCellStyle(style);
            // 创建新文件
            File dir = new File(TEMPLATE_FILE_DIR);
            if (!dir.exists()) {
                dir.mkdir();
            }
            File file = new File(TEMPLATE_FILE_DIR, TEMPLATE_FLE_NAME + TEMPLATE_FLE_SUFFIX);
            // 保存到磁盘
            SimpleExcelUtils.write(workbook,TEMPLATE_FLE_NAME,null,new FileOutputStream(file));
        }
        // 下载
       SimpleExcelUtils.write(workbook,TEMPLATE_FLE_NAME,response,null);

    }

    public void exportData(HttpServletResponse response) throws IOException {
        // 第一步，创建一个webbook，对应一个Excel文件
        XSSFWorkbook workbook = new XSSFWorkbook();
        // 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
        XSSFSheet sheet = workbook.createSheet("Sheet1");
        // 设置列宽度自适应
        for (int i = 0; i < COLUMN_NAME.length; i++) {
            sheet.autoSizeColumn(i, true);
        }

        // 第三步，设置标题列
        XSSFRow row = sheet.createRow(0);
        // 第四步，创建单元格，并设置值表头 设置表头居中
        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER); // 创建一个居中格式
        XSSFCell cell = row.createCell(0);
        cell.setCellValue(COLUMN_NAME[0]);
        cell.setCellStyle(style);
        cell = row.createCell(1);
        cell.setCellValue(COLUMN_NAME[1]);
        cell.setCellStyle(style);
        cell = row.createCell(2);
        cell.setCellValue(COLUMN_NAME[2]);
        cell.setCellStyle(style);
        cell = row.createCell(3);
        cell.setCellValue(COLUMN_NAME[3]);
        cell.setCellStyle(style);

        // 第五步，写入实体数据 实际应用中这些数据从数据库得到，
        List<Shop> list = new ArrayList<>(shopList);

        for (int i = 0; i < list.size(); i++) {
            row = sheet.createRow(i + 1);
            Shop shop = list.get(i);
            // 第四步，创建单元格，并设置值
            row.createCell(0).setCellValue(shop.getShopName());
            row.createCell(1).setCellValue(shop.getShopId());
            cell = row.createCell(2);
//            cell.setCellValue(new SimpleDateFormat("yyyy-MM-dd").format(stu.getTime()));
            cell.setCellValue(new SimpleDateFormat("HH:mm:ss").format(shop.getStartTime()));
            cell = row.createCell(3);
            cell.setCellValue(new SimpleDateFormat("HH:mm:ss").format(shop.getEndTime()));
        }
        //第六步,输出Excel文件
        try {
            SimpleExcelUtils.write(workbook,TEMPLATE_FLE_NAME,response,null);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 导入excel文件
     * 1. 根据文件类型创建 Workbook
     * 2. 根据 Workbook 创建 Sheet
     * 3. 根据 Sheet 获取 Row
     * 4. 根据 Row 获取 Cell,并将 Cell 的数据设置到对象对应属性中
     * 5. 入库
     *
     * @param file
     */
    public void importData(MultipartFile file) throws Exception {
        long start = System.currentTimeMillis();

        try {
            // 根据流创建 Workbook，需要根据文件类型，确定 Workbook 实现类
            Workbook workbook = SimpleExcelUtils.createWorkbook(file);
            if (workbook == null) {
                throw new RuntimeException("创建 Workbook 失败");
            }
            // 创建 Sheet
            Sheet sheet = workbook.getSheet("Sheet1");
            int rows = sheet.getLastRowNum();// 指的行数，一共有多少行:从0开始
            if (rows == 0) {
                throw new Exception("模板数据不能为空");
            }
            // 待入库的门店数据
            // 从第一行开始遍历
            for (int i = 1; i <= rows + 1; i++) {
                Row row = sheet.getRow(i);
                // 如果行不为空，读取行中的每一列的数据，并设置到对象中
                if (row != null) {
                    Shop shop = new Shop();
                    // 门店名称
                    String shopName = SimpleExcelUtils.getCellValue(row.getCell(0));
                    if (shopName == null || shopName.trim().length() <= 0) {
                        log.error("门店名称非法：第" + (i + 1) + "行");
                        return;
                    }
                    shop.setShopName(shopName);
                    // 门店编号
                    shop.setShopId(Long.valueOf(SimpleExcelUtils.getCellValue(row.getCell(1))));
                    // 开始时间
                    String startTime = SimpleExcelUtils.getCellValue(row.getCell(2));
                    SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                    shop.setStartTime(dateFormat.parse(startTime));
                    // 结束时间
                    String endTime = SimpleExcelUtils.getCellValue(row.getCell(3));
                    shop.setEndTime(dateFormat.parse(endTime));
                    log.info("保存门店:" + shop.toString());
                    shopList.add(shop);
                }
            }

            log.info("入库:" + shopList.toString());
        } finally {
            long end = System.currentTimeMillis();

            System.out.println("耗时：" + (end - start));
        }


    }


    /**
     * 导入excel文件
     * 1. 根据文件类型创建 Workbook
     * 2. 根据 Workbook 创建 Sheet
     * 3. 根据 Sheet 获取 Row
     * 4. 根据 Row 获取 Cell,并将 Cell 的数据设置到对象对应属性中
     * 5. 入库
     *
     * @param file
     */
    public void importDataByThread(MultipartFile file) throws Exception {
        long start = System.currentTimeMillis();

        try {
            // 根据流创建 Workbook，需要根据文件类型，确定 Workbook 实现类
            Workbook workbook = SimpleExcelUtils.createWorkbook(file);
            if (workbook == null) {
                throw new RuntimeException("创建 Workbook 失败");
            }
            // 创建 Sheet
            Sheet sheet = workbook.getSheet("Sheet1");
            int rows = sheet.getLastRowNum();// 指的行数，一共有多少行:从0开始
            if (rows == 0) {
                throw new Exception("模板数据不能为空");
            }
            int worker = (int) Math.ceil((double) rows / 2000);

            if (worker > 1) {
                ThreadPoolExecutor threadPool = null;
                try {
                    List<Future<ParseResult<Shop>>> futures = new ArrayList<>();
                    threadPool = new ThreadPoolExecutor(worker, worker + 1,
                            1000L, TimeUnit.SECONDS, new LinkedBlockingDeque<>());
                    for (int i = 0; i < worker; i++) {
                        int startIndex = 1 + 2000 * i;
                        int endIndex = Math.min(2000 * (i + 1), rows);
                        Task task = new Task(startIndex, endIndex, sheet);
                        Future<ParseResult<Shop>> future = threadPool.submit(task);
                        futures.add(future);
                    }
                    List<Shop> shops = new ArrayList<>();
                    for (Future<ParseResult<Shop>> future : futures) {
                        // 阻塞主线程，获取子线程执行结果
                        ParseResult<Shop> shopParseResult = future.get();
                        log.info("线程执行结果:" + shopParseResult);
                        if (!shopParseResult.getSuccess()){
                            // 存在非法数据
                            // TODO 如何在一个线程发现数据非法时，通知其他线程停止执行呢
                            log.error(shopParseResult.getErrMsg());
                            return;
                        }
                        // 将结果汇总
                        shops.addAll(shopParseResult.getData());
                    }
                    log.info("文件上次校验通过，开始入库:"+ shops.size());
                    // TODO 直接调用dao的批量插入接口，还是再使用线程呢

                } finally {
                    threadPool.shutdown();
                }

            } else {
                // 待入库的门店数据
                // 从第一行开始遍历
                for (int i = 1; i <= rows + 1; i++) {
                    Row row = sheet.getRow(i);
                    // 如果行不为空，读取行中的每一列的数据，并设置到对象中
                    if (row != null) {
                        Shop shop = new Shop();
                        // 门店名称
                        String shopName = SimpleExcelUtils.getCellValue(row.getCell(0));
                        if (shopName == null || shopName.trim().length() <= 0) {
                            log.error("门店名称非法：第" + (i + 1) + "行");
                            return;
                        }
                        shop.setShopName(shopName);
                        // 门店编号
                        shop.setShopId(Long.valueOf(SimpleExcelUtils.getCellValue(row.getCell(1))));
                        // 开始时间
                        String startTime = SimpleExcelUtils.getCellValue(row.getCell(2));
                        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                        shop.setStartTime(dateFormat.parse(startTime));
                        // 结束时间
                        String endTime = SimpleExcelUtils.getCellValue(row.getCell(3));
                        shop.setEndTime(dateFormat.parse(endTime));
                        log.info("保存门店:" + shop.toString());
                        shopList.add(shop);
                    }
                }

                log.info("入库:" + shopList.toString());
            }
        } finally {
            long end = System.currentTimeMillis();

            System.out.println("耗时：" + (end - start));
        }
    }
}
