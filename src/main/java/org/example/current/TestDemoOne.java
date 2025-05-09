package org.example.current;

import lombok.extern.slf4j.Slf4j;

import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * Create By ecchen
 * Date 2025/4/24 16:12
 * Description
 */
@Slf4j
public class TestDemoOne {
    public static void main(String[] args) throws Exception {
        concurrenceTest();
    }

    public static void concurrenceTest() {
        /**
         * 模拟高并发情况代码
         */
        final AtomicInteger atomicInteger = new AtomicInteger(0);
        final CountDownLatch countDownLatch = new CountDownLatch(1000); // 相当于计数器，当所有都准备好了，再一起执行，模仿多并发，保证并发量
        final CountDownLatch countDownLatch2 = new CountDownLatch(1000); // 保证所有线程执行完了再打印atomicInteger的值
        ExecutorService executorService = Executors.newFixedThreadPool(10);
        try {
            for (int i = 0; i < 1000; i++) {
                executorService.submit(() -> {
                    try {
                        countDownLatch.await(); //一直阻塞当前线程，直到计时器的值为0,保证同时并发
                    } catch (InterruptedException e) {
                        log.error(e.getMessage(), e);
                    }
                    //每个线程增加1000次，每次加1
                    for (int j = 0; j < 1000; j++) {
                        atomicInteger.incrementAndGet();
                    }
                    countDownLatch2.countDown();
                });
                countDownLatch.countDown();
            }

            countDownLatch2.await();// 保证所有线程执行完
            executorService.shutdown();
            log.info("atomicInteger的值为：{}", atomicInteger.get());
        } catch (Exception e) {
            log.error(e.getMessage(), e);
        }
    }
}
