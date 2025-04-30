package org.example.controller;

import cn.hutool.core.io.FileUtil;
import cn.hutool.core.io.resource.ClassPathResource;
import groovy.lang.Binding;
import groovy.lang.GroovyClassLoader;
import groovy.lang.GroovyShell;
import groovy.lang.Script;
import lombok.extern.slf4j.Slf4j;
import org.codehaus.groovy.runtime.InvokerHelper;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.nio.charset.StandardCharsets;
import java.util.HashMap;
import java.util.Map;

/**
 * Create By ecchen
 * Date 2025/4/25 16:41
 * Description
 */
@RestController
@RequestMapping("/groovy")
@Slf4j
public class GroovyController {

    /**
     * 缓存Script，避免创建太多
     */
    private static final Map<String, Script> SCRIPT_MAP = new HashMap<>();

    private static final GroovyClassLoader CLASS_LOADER = new GroovyClassLoader();

    @RequestMapping("/test")
    public String test() {
        //创建GroovyShell
//        GroovyShell groovyShell = new GroovyShell();
        ClassPathResource classPathResource = new ClassPathResource("groovy/BeanCase.groovy");
        String scriptText = FileUtil.readString(classPathResource.getFile(), StandardCharsets.UTF_8);
        Script script = loadScript("beanCase", scriptText);
        //装载解析脚本代码
//        Script script = groovyShell.parse(scriptText);
        //执行
        script.invokeMethod("getBean", null);
        return "ok";
    }

    /**
     * 每次调用这个方法都创建了GroovyShell、Script等实例，随着调用次数的增加，必然会出现OOM。缓存Script，避免创建太多，优化方法，避免OOM，但是如果脚本内容修改，需要情况缓存，重新装载脚本实例
     * @param key
     * @param rule
     * @return
     */
    public static Script loadScript(String key, String rule) {
        if (SCRIPT_MAP.containsKey(key)) {
            return SCRIPT_MAP.get(key);
        }
        Script script = loadScript(rule, new Binding());
        SCRIPT_MAP.put(key, script);
        return script;
    }


    public static Script loadScript(String rule, Binding binding) {
        if (StringUtils.isEmpty(rule)) {
            return null;
        }
        try {
            Class ruleClazz = CLASS_LOADER.parseClass(rule);
            if (ruleClazz != null) {
                log.info("load rule:" + rule + " success!");
                return InvokerHelper.createScript(ruleClazz, binding);
            }
        } catch (Exception e) {
            log.error(e.getMessage(), e);
        } finally {
            CLASS_LOADER.clearCache();
        }
        return null;
    }
}
