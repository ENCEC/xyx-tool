package org.example;

import lombok.extern.slf4j.Slf4j;
import org.springframework.boot.context.event.ApplicationReadyEvent;
import org.springframework.context.ApplicationListener;
import org.springframework.stereotype.Component;

/**
 * Create By ecchen
 * Date 2025/10/9 16:03
 * Description
 */
@Component
@Slf4j
public class ApplicationStartListener implements ApplicationListener<ApplicationReadyEvent> {
    @Override
    public void onApplicationEvent(ApplicationReadyEvent applicationReadyEvent) {
        // 这里可以做些启动初始化操作，此时应用bean已经全部注册完，代表启动成功，整个应用都OK了，可以对外服务了
        // 与postConstruct区别的在于，postConstruct只能保证当前bean注册完，可能依赖的bean还没注册完
        log.info("====xyx应用启动成功====");
    }
 /**
            | 维度              | `@PostConstruct`           | `ApplicationReadyEvent`                             |
            | --------------- | -------------------------- | --------------------------------------------------- |
            | **所属阶段**        | Bean 初始化阶段（Post-Process）   | SpringApplication 启动完毕阶段                            |
            | **发布时机**        | 所在 Bean 被完全装配后立即执行（依赖注入结束） | 所有 Bean、回调、CommandLineRunner、ApplicationRunner 都跑完后 |
            | **是否已起 web 服务** | ❌ 尚未启动                     | ✅ 内嵌 Tomcat/Jetty/Netty 已监听端口                       |
            | **事务是否可用**      | ❌ 可能尚未建连接池                 | ✅ 数据源、事务、JPA、MQ 等全部就绪                               |
            | **执行线程**        | 创建该 Bean 的线程（同步阻塞）         | 事件广播线程（可配异步）                                        |
            | **失败影响**        | Bean 创建失败 → 整个上下文启动失败      | 监听器抛异常不会阻止应用已“UP”                                   |
            | **顺序控制**        | 无法保证依赖其他 Bean 的初始化顺序       | 可借助 `@Order`/`SmartInitializingSingleton` 或额外事件     |
            | **典型用途**        | 校验必填配置、预热缓存、启动定时器          | 发送启动成功通知、健康检查、优雅注册服务、开启流量                           |
**/
}
