package org.example.order;

import lombok.AllArgsConstructor;
import lombok.Data;

/**
 * Create By ecchen
 * Date 2025/10/9 16:36
 * Description
 */
@Data
@AllArgsConstructor
public class MsgEvent {

    /** 该类型事件携带的信息 */
    public String orderId;
}
