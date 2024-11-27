package org.example;

import lombok.Data;

import java.math.BigDecimal;

/**
 * @Auther: chenec
 * @Date: 2024/2/17 14:58
 * @Description: ShopInfo
 * @Version 1.0.0
 */
@Data
public class ShopInfo {
    /**
     * 店铺名称
     */
    private String shopName;
    /**
     * 订单编号
     */
    private String orderNo;
    /**
     * 价格
     */
    private BigDecimal price;
}
