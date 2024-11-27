package org.example.dto;

import lombok.Data;

import java.math.BigDecimal;

/**
 * @Auther: chenec
 * @Date: 2024/2/20 16:24
 * @Description: FurnitureLogisticDto
 * @Version 1.0.0
 */
@Data
public class FurnitureLogisticDto {
    /**
     * 物流单号
     */
    private String logisticNo;
    /**
     * 店铺
     */
    private String shopName;
    /**
     * 订单号
     */
    private String orderNo;
    /**
     * 金额
     */
    private BigDecimal fcy;
}
