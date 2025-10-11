package org.example;

import org.example.order.OrderService;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

/**
 * Create By ecchen
 * Date 2025/10/9 16:23
 * Description
 */
@SpringBootTest
public class ApplicationTest {
    @Autowired
    private OrderService orderService;
    @Test
    public void contextLoads() {
        orderService.buyOrder("561058");
    }

    @Test
    public void buyOrderTest() {
        orderService.buyOrder("732171109");
    }
}
