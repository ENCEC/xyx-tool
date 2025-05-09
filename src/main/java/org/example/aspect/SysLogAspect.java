package org.example.aspect;

import cn.hutool.core.convert.Convert;
import lombok.extern.slf4j.Slf4j;
import org.aspectj.lang.ProceedingJoinPoint;
import org.aspectj.lang.annotation.Around;
import org.aspectj.lang.annotation.Aspect;
import org.aspectj.lang.annotation.Pointcut;
import org.springframework.core.Ordered;
import org.springframework.core.annotation.Order;
import org.springframework.stereotype.Component;

/**
 * Create By ecchen
 * Date 2025/4/30 10:43
 * Description
 */
@Component
@Aspect
@Order(Ordered.HIGHEST_PRECEDENCE)
@Slf4j
public class SysLogAspect {

    @Pointcut("@annotation(org.example.annotion.SysLog)")
    public void pointCut() {

    }

    @Around("pointCut()")
    public Object around(ProceedingJoinPoint point) throws Throwable {
        //逻辑开始时间
        long beginTime = System.currentTimeMillis();
        log.info("====执行前====");
        //执行方法
        Object result = point.proceed();
        log.info("====执行后====");
        //todo，保存日志，自己完善
        saveLog(point, beginTime);

        return result;
    }

    private void saveLog(ProceedingJoinPoint point, long beginTime) {
        log.info("执行时间：" + Convert.toDate(beginTime));
    }
}
