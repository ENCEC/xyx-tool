package org.example.config;

import org.example.aspect.RepeatSubmitInterceptor;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Configuration;
import org.springframework.web.servlet.config.annotation.InterceptorRegistry;
import org.springframework.web.servlet.config.annotation.WebMvcConfigurer;

/**
 * Create By ecchen
 * Date 2025/4/30 11:10
 * Description
 */
@Configuration
public class WebConfig implements WebMvcConfigurer {

    @Autowired
    private RepeatSubmitInterceptor repeatSubmitInterceptor;

    @Override
    public void addInterceptors(InterceptorRegistry registry) {
        //不拦截的uri
        final String[] commonExclude = {};
        registry.addInterceptor(repeatSubmitInterceptor).excludePathPatterns(commonExclude);
    }
}