package org.example.handler;

import lombok.extern.slf4j.Slf4j;
import org.springframework.web.bind.annotation.ControllerAdvice;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.servlet.ModelAndView;

import java.io.IOException;

/**
 * @Auther: chenec
 * @Date: 2024/3/9 16:41
 * @Description: MyExceptionHandler
 * @Version 1.0.0
 */
@ControllerAdvice
@Slf4j
public class MyExceptionHandler {

    @ExceptionHandler(value = IOException.class)
    public ModelAndView exceptionHandler(IOException e) {
        log.error("全局异常捕获>>>文件流异常:{}", e);
        ModelAndView modelAndView = new ModelAndView();
        modelAndView.addObject("exception", "读取文件流异常，请联系管理员！");
        modelAndView.setViewName("error");
        return modelAndView;
    }

    @ExceptionHandler(value = Exception.class)
    public ModelAndView exceptionHandler(Exception e) {
        log.error("全局异常捕获>>>:{}", e);
        ModelAndView modelAndView = new ModelAndView();
        modelAndView.addObject("exception", e.getMessage());
        modelAndView.setViewName("error");
        return modelAndView;
    }
}
