package com.ooooor.exl4j.model;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;

public class JoomOrder extends BaseRowModel {

    @ExcelProperty(index = 1)
    private String orderId;

    @ExcelProperty(index = 14)
    private String extraPrice;

    public String getOrderId() {
        return orderId;
    }

    public void setOrderId(String orderId) {
        this.orderId = orderId;
    }

    public String getExtraPrice() {
        return extraPrice;
    }

    public void setExtraPrice(String extraPrice) {
        this.extraPrice = extraPrice;
    }

    @Override
    public String toString() {
        return "JoomOrder{" +
                "orderId='" + orderId + '\'' +
                ", extraPrice='" + extraPrice + '\'' +
                '}';
    }
}
