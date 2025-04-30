package groovy

import org.example.config.GroovyTestService
import org.example.config.SpringContextUtil

/**
 * 静态变量
 */
class Globals {
    static String PARAM1 = "静态变量"
    static int[] arrayList = [1, 2]
}

def getBean() {
    GroovyTestService groovyTestService = SpringContextUtil.getBean(GroovyTestService.class);
    groovyTestService.test()
}