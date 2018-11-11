package com.ooooor.xls4j.application.dto;

import lombok.Data;

/**
 * @description:
 * @author: chenr
 * @date: 18-11-9
 */
@Data
public class AjaxRes {
    private Integer code;
    private String msg;
    private Object obj;

    public void setSucceed(Object obj){
        this.obj = obj;
        this.setCode(1);
    }
}
