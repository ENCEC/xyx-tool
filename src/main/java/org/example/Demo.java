package org.example;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ApplicationEvent;
import org.springframework.context.ApplicationListener;
import org.springframework.context.ConfigurableApplicationContext;
import org.springframework.context.annotation.Bean;
import org.springframework.context.event.ApplicationEventMulticaster;
import org.springframework.context.event.SimpleApplicationEventMulticaster;
import org.springframework.context.support.AbstractApplicationContext;
import org.springframework.scheduling.concurrent.ThreadPoolTaskExecutor;
import org.springframework.stereotype.Component;

/**
 * Create By ecchen
 * Date 2025/10/10 10:41
 * Description
 */
@SpringBootApplication
public class Demo {
    public static void main(String[] args) {
        ConfigurableApplicationContext ctx =
                SpringApplication.run(Demo.class, args);

        // 发布事件
        ctx.publishEvent(new MyEvent("hello"));
    }

    static class MyEvent extends ApplicationEvent {
        public MyEvent(Object source) { super(source); }
    }

    @Component
    static class MyListener implements ApplicationListener<MyEvent> {
        @Override
        public void onApplicationEvent(MyEvent event) {
            System.out.println("收到：" + event.getSource());
        }
    }

    // 把广播器换成异步的
    @Bean(name = AbstractApplicationContext.APPLICATION_EVENT_MULTICASTER_BEAN_NAME)
    public ApplicationEventMulticaster multicaster() {
        SimpleApplicationEventMulticaster caster =
                new SimpleApplicationEventMulticaster();
        ThreadPoolTaskExecutor exec = new ThreadPoolTaskExecutor();
        exec.setCorePoolSize(4);
        exec.initialize();
        caster.setTaskExecutor(exec);
        return caster;
    }
}