package org.example.dto;

import lombok.Data;

import java.math.BigDecimal;

/**
 * @Auther: chenec
 * @Date: 2024/2/20 16:38
 * @Description: FurnitureSpecDto
 * @Version 1.0.0
 */
@Data
public class FurnitureSpecDto {
    /**
     * 规格
     */
    private String spec;
    /**
     * 货号
     */
    private String productNo;
    /**
     * 成本
     */
    private BigDecimal cost;
}
