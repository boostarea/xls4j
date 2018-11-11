package com.ooooor.xls4j.application.dto;

import lombok.Data;

import java.io.Serializable;
import java.util.List;


@Data
public class OneLineResultDto implements Serializable{
    /**行索引*/
    private Long lineNum;
    /**行数据*/
    private List<Object> lineData;
    /**异常信息*/
    private String errMsg;
}
