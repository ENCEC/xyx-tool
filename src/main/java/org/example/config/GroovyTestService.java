package org.example.config;
import org.springframework.stereotype.Service;

/**
 * Create By ecchen
 * Date 2025/4/25 16:36
 * Description
 */

@Service
public class GroovyTestService {

    public void test(){
        System.out.println("我是SpringBoot框架的成员类，但该方法由Groovy脚本调用");
    }

}
