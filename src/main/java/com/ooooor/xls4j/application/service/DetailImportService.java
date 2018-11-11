package com.ooooor.xls4j.application.service;

import com.ooooor.xls4j.application.dto.OneLineResultDto;

import java.util.List;
import java.util.Map;


public interface DetailImportService {

    /**
     * 处理excel表格的一行记录
     * @param headerIndexMap
     * @param currentRowData
     * @param currentRowNum
     * @param importTrxId
     * @return
     */
    OneLineResultDto doImport(Map<String, Integer> headerIndexMap, List<Object> currentRowData, Long currentRowNum, String importTrxId);

    /**
     * 获取上传类型
     * @return
     */
    default Byte getOperateType() {
        return 0;
    }
}
