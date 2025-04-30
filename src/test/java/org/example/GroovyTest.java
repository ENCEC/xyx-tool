package org.example;

import cn.hutool.core.io.FileUtil;
import cn.hutool.core.io.resource.ClassPathResource;
import groovy.lang.GroovyShell;
import groovy.lang.Script;

import java.nio.charset.StandardCharsets;

/**
 * Create By ecchen
 * Date 2025/4/25 16:20
 * Description
 */
public class GroovyTest {
    public static void main(String[] args) throws Exception {
        //创建GroovyShell
        GroovyShell groovyShell = new GroovyShell();
        //装载解析脚本代码
        ClassPathResource classPathResource = new ClassPathResource("groovy/HelloWorld.groovy");
        String scriptText = FileUtil.readString(classPathResource.getFile(), StandardCharsets.UTF_8);
        Script script = groovyShell.parse(scriptText);
        //执行
        script.invokeMethod("HelloWorld", null);
    }
}
