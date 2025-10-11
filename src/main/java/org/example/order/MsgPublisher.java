package org.example.order;

import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.ApplicationContext;
import org.springframework.stereotype.Component;

/**
 * Create By ecchen
 * Date 2025/10/9 16:38
 * Description
 */
@Component
@Slf4j
public class MsgPublisher {
    @Autowired
    private ApplicationContext applicationContext;
    /**
     * 下单
     * @param orderId 订单ID
     */
    public String buyOrder(String orderId) {
        long start = System.currentTimeMillis();
        // 1.查询订单详情

        // 2.检验订单价格 （同步处理）
        applicationContext.publishEvent(new OrderProductEvent(this, orderId));

        // 3.短信通知（异步处理）
        applicationContext.publishEvent(new MsgEvent(orderId));

        long end = System.currentTimeMillis();
        log.info("任务全部完成，总耗时：({})毫秒", end - start);
        return "购买成功";
    }
}
