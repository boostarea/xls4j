package com.ooooor.xls4j.application.service;

import com.ooooor.xls4j.application.dto.OneLineResultDto;
import com.ooooor.xls4j.domain.exception.ServiceException;
import com.ooooor.xls4j.infrastructure.controller.acceptor.FileAcceptorController;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.redis.core.RedisTemplate;
import org.springframework.stereotype.Service;

import java.util.List;
import java.util.Map;

/**
 * @Author: chenr
 * @Date: Create in 18-2-11 上午10:20
 */
@Service
public class OutOrderImportServiceImpl implements DetailImportService {

    private Logger logger = LoggerFactory.getLogger(OutOrderImportServiceImpl.class);

    @Autowired
    protected RedisTemplate<String, String> redisTemplate;

    @Override
    public OneLineResultDto doImport(Map<String, Integer> headerIndexMap, List<Object> currentRowData, Long currentRowNum,
                                     String importTrxId) {
        try {
            headerIndexMap.keySet();
            redisTemplate.opsForHash().increment(importTrxId, FileAcceptorController.HASH_FIELD_IMPORT_SUCCESS_COUNT,
                    1L);
            // currentRowData.get(headerIndexMap.get("出库单号"));
        } catch (Exception e) {
            logger.error("导入第" + currentRowNum + "行物流单号" + currentRowData + "异常：" + e.getMessage(), e);
            return getOneLineResultDto(currentRowData, currentRowNum, getExceptionMsg(e));
        }
        return null;
    }

    private String getExceptionMsg(Exception e) {
        String errMsg = "系统异常，请联系系统管理员";
        if (e instanceof ServiceException) {
            errMsg = e.getMessage();
        }
        return errMsg;
    }

    private OneLineResultDto getOneLineResultDto(List<Object> currentRowData, Long currentRowNum, String message) {
        OneLineResultDto resultDto = new OneLineResultDto();
        resultDto.setErrMsg(message);
        resultDto.setLineData(currentRowData);
        resultDto.setLineNum(currentRowNum);
        return resultDto;
    }

}
