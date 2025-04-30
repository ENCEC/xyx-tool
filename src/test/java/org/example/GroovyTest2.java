package org.example;

import cn.hutool.core.io.FileUtil;
import cn.hutool.core.io.resource.ClassPathResource;
import groovy.lang.GroovyShell;
import groovy.lang.Script;

import java.nio.charset.StandardCharsets;
import java.util.HashMap;
import java.util.Map;

/**
 * Create By ecchen
 * Date 2025/4/25 16:20
 * Description
 */
public class GroovyTest2 {
    public static void main(String[] args) throws Exception {
        //创建GroovyShell
        GroovyShell groovyShell = new GroovyShell();
        //装载解析脚本代码
        ClassPathResource classPathResource = new ClassPathResource("groovy/AddCase.groovy");
        String scriptText = FileUtil.readString(classPathResource.getFile(), StandardCharsets.UTF_8);
        Script script = groovyShell.parse(scriptText);

        //执行加法脚本
        Object[] params1 = new Object[]{1, 2};
        int sum = (int) script.invokeMethod("add", params1);
        System.out.println("a加b的和为:" + sum);
        //执行解析脚本
        Map<String, String> paramMap = new HashMap<>();
        paramMap.put("科目1", "语文");
        paramMap.put("科目2", "数学");
        Object[] params2 = new Object[]{paramMap};
        String result = (String) script.invokeMethod("mapToString", params2);
        System.out.println("mapToString:" + result);
    }
}
