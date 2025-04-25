package org.example;

import cn.hutool.core.io.FileUtil;
import cn.hutool.core.io.resource.ClassPathResource;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
//import org.junit.Test;
//import org.junit.runner.RunWith;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.data.redis.core.StringRedisTemplate;
import org.springframework.data.redis.core.script.DefaultRedisScript;
import org.springframework.data.redis.core.script.RedisScript;
import org.springframework.test.context.junit4.SpringRunner;

import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;

/**
 * Create By ecchen
 * Date 2025/4/3 11:16
 * Description
 */
@SpringBootTest(classes = Application.class)
//@RunWith(SpringRunner.class)
@Slf4j
public class LuaDemo {

    @Autowired
    private StringRedisTemplate redisTemplate;

    @Test
    @SneakyThrows
    public void testScriptLoad() {
//        String luaScript = "local num = redis.call('incr', KEYS[1])\n" +
//                "if tonumber(num) == 1 then\n" +
//                "\tredis.call('expire', KEYS[1], ARGV[1])\n" +
//                "\treturn 1\n" +
//                "elseif tonumber(num) > tonumber(ARGV[2]) then\n" +
//                "\treturn 0\n" +
//                "else \n" +
//                "\treturn 1\n" +
//                "end\n";
        ClassPathResource resource = new ClassPathResource("lua/rate_limit.lua");
        String luaScript = FileUtil.readString(resource.getFile(), StandardCharsets.UTF_8);
        RedisScript<Long> script = new DefaultRedisScript<>(luaScript, Long.class);

        // 构造参数
        String key = "rate_limit:test_key";    // KEYS[1]
        Integer expireTime = 600;               // ARGV[1]（过期时间，单位秒）
        Integer threshold = 3;                 // ARGV[2]（阈值）

        // 执行脚本（直接返回 0 或 1）
        Long result = redisTemplate.execute(
                script,
                Collections.singletonList(key),   // KEYS 列表
                expireTime.toString(),             // ARGV[1]
                threshold.toString()              // ARGV[2]
        );

        System.out.println("执行结果: " + result); // 输出 0 或 1
    }

    @Test
    public void testAddLua() {
        ClassPathResource resource = new ClassPathResource("lua/simple_add.lua");
        String luaScript = FileUtil.readString(resource.getFile(), StandardCharsets.UTF_8);
        RedisScript<Long> script = new DefaultRedisScript<>(luaScript, Long.class);
        Long result = redisTemplate.execute(script, new ArrayList<>(), "10","20");
        System.out.println("执行结果: " + result);
    }
}
