package org.example.order;

import lombok.*;
import org.springframework.context.ApplicationEvent;

/**
 * Create By ecchen
 * Date 2025/10/9 16:11
 * Description
 */
@Getter
@ToString
public class OrderProductEvent extends ApplicationEvent {
    /** 该类型事件携带的信息 */
    private String orderId;
    public OrderProductEvent(Object source){
        super(source);
    }
    public OrderProductEvent(Object source, String orderId) {
        super(source);
        this.orderId = orderId;
    }
}
