package com.wxss.threadexcel.controller;
import com.wxss.threadexcel.domain.vo.ResultVO;

import com.wxss.threadexcel.service.ShopService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;

@Controller
@RequestMapping("shop")
public class ShopController {

    @Autowired
    private ShopService shopService;

    /**
     * 下载模板
     * @return
     */
    @GetMapping("downloadTemplate")
    @ResponseBody
    public void downloadTemplate(HttpServletResponse response) throws Exception {
        shopService.downloadTemplate(response);
    }
    @GetMapping("export")
    @ResponseBody
    public void exportData(HttpServletResponse response) throws Exception {
        shopService.exportData(response);
    }

    @GetMapping("import")
    @ResponseBody
    public ResultVO<Object> importData(MultipartFile file) throws Exception {
        shopService.importDataByThread(file);
//        shopService.importData(file);
        return ResultVO.success();
    }

}
