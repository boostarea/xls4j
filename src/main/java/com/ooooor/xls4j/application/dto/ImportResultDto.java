package com.ooooor.xls4j.application.dto;

import java.util.List;
import java.util.Map;


public class ImportResultDto {
    private Map<String, Integer> headerIndexMap;

    List<OneLineResultDto> errorLines;

    private Long successCount;

    public ImportResultDto() {
    }

    public ImportResultDto(Map<String, Integer> headerIndexMap, List<OneLineResultDto> errorLines, Long successCount) {
        this.headerIndexMap = headerIndexMap;
        this.errorLines = errorLines;
        this.successCount = successCount;
    }

    public Long getSuccessCount() {
        return successCount;
    }

    public void setSuccessCount(Long successCount) {
        this.successCount = successCount;
    }

    public Map<String, Integer> getHeaderIndexMap() {
        return headerIndexMap;
    }

    public void setHeaderIndexMap(Map<String, Integer> headerIndexMap) {
        this.headerIndexMap = headerIndexMap;
    }

    public List<OneLineResultDto> getErrorLines() {
        return errorLines;
    }

    public void setErrorLines(List<OneLineResultDto> errorLines) {
        this.errorLines = errorLines;
    }
}
