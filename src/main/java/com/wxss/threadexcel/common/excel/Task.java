package com.wxss.threadexcel.common.excel;

import com.wxss.threadexcel.domain.entity.Shop;
import com.wxss.threadexcel.utils.SimpleExcelUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.Callable;

public class Task implements Callable<ParseResult<Shop>> {

    private Integer start;
    private Integer end;
    private Sheet sheet;

    public Task(Integer start, Integer end, Sheet sheet) {
        this.start = start;
        this.end = end;
        this.sheet = sheet;
    }


    @Override
    public ParseResult<Shop> call() throws Exception {

        ParseResult<Shop> result = new ParseResult<Shop>();


        List<Shop> shops = new ArrayList<>();
        System.err.println(Thread.currentThread().getName() + "开始执行 start"+ start + " ;end" + end);

        for (int i = start; i <= end; i++) {
            Row row = sheet.getRow(i);
            // 如果行不为空，读取行中的每一列的数据，并设置到对象中
            if (row != null) {
                Shop shop = new Shop();
                // 门店名称
                String shopName = SimpleExcelUtils.getCellValue(row.getCell(0));
                if (shopName == null || shopName.trim().length() <= 0) {
                    System.err.println(Thread.currentThread().getName() + "start"+ start + " ;end" + end+ ";门店名称非法：第" + (i + 1) + "行");
                    result.setSuccess(false);
                    result.setErrMsg(Thread.currentThread().getName() + "start"+ start + " ;end" + end+ ";门店名称非法：第" + (i + 1) + "行");

                    return result;
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
                shops.add(shop);
            }
        }

        System.out.println(Thread.currentThread().getName() + "保存门店成功: start"+ start + " ;end" + end);

        result.setSuccess(true);
        result.setData(shops);
        return result;
    }
}
