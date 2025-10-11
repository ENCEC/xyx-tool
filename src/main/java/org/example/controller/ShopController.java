package org.example.controller;

import lombok.extern.slf4j.Slf4j;
import org.example.annotion.RepeatSubmit;
import org.example.annotion.SysLog;
import org.example.service.ShopService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;

/**
 * @Auther: chenec
 * @Date: 2024/2/18 17:51
 * @Description: ShopController
 * @Version 1.0.0
 */
@Controller
@RequestMapping("/shop")
@Slf4j
public class ShopController {
    @Autowired
    private ShopService shopService;

    @GetMapping("/test")
    @ResponseBody
    @SysLog
    @RepeatSubmit
    public String test() {
        log.info("execute test method");
        return "hello test content";
    }

    @GetMapping("/gotoUploadPage")
    public String gotoUploadPage() {
        return "uploadPage";
    }

    @PostMapping("/generatorScreenDayReport")
    public String generatorScreenDayReport(@RequestPart("file") MultipartFile multipartFile, HttpServletResponse response) throws Exception {
        shopService.generatorScreenDayReport(multipartFile, response);
        return "success";
    }

    @PostMapping("/generatorScreenMonthReport")
    public void generatorMonthReport(@RequestPart("file") MultipartFile multipartFile, HttpServletResponse response) throws Exception {
        shopService.generatorScreenMonthReport(multipartFile, response);
    }
}
