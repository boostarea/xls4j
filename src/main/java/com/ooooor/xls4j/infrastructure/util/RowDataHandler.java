package com.ooooor.xls4j.infrastructure.util;

/**
 * @description:
 * @author: chenr
 * @date: 18-11-11
 */
public interface RowDataHandler {
/*
    *//**
     * 表头: Map<列名,列索引>
     *//*
    Map<String, Integer> headerIndexMap = Maps.newHashMap();

    *//**
     * 导入失败返回的信息
     *//*
    List<OneLineResultDto> errorLines = Lists.newArrayList();
    *//**
     * 总行数
     *//*
    long lineCount = 0;

    long getLineCount();

    Object[] getCurrentRowData();

    DetailImportService getDetailImportService();

    Map<String, Integer> getHeaderIndexMap();

    Map<Long, OneLineResultDto> getErrorResultHolder();

    RedisTemplate<String, String> getRedisTemplate();

    String getImportTrxId();

    ExecutorService getThreadPool();

    default void handleRowData() {
        if (getLineCount() == 1) {
            for (int i = 0; i < getCurrentRowData().length; i++) {
                if (getCurrentRowData()[i] == null) {
                    break;
                }

                getHeaderIndexMap().put(getCurrentRowData()[i].toString(), i);
            }
            getRedisTemplate().opsForHash().increment(getImportTrxId(), AbstractExcelImportController.HASH_FIELD_EXCEL_SUCCESS_COUNT, 1L);
            return;
        }
        long currentLineCount = getLineCount();
        List<Object> list = Lists.newArrayList(Arrays.copyOf(getCurrentRowData(), getHeaderIndexMap().size()));

        Thread thread =
                new Thread() {
                    UserInfo userInfo = UserContext.getCurrentUser();

                    @Override
                    public void run() {
                        UserContext.putUserInfo(userInfo);
                        //处理一行记录
                        OneLineResultDto result = getDetailImportService().doImport(getHeaderIndexMap(), list, currentLineCount, getImportTrxId());

                        if (result == null) {
                            getRedisTemplate().opsForHash().increment(getImportTrxId(), AbstractExcelImportController.HASH_FIELD_EXCEL_SUCCESS_COUNT, 1L);
                        } else {
                            getRedisTemplate().opsForHash().increment(getImportTrxId(), AbstractExcelImportController.HASH_FIELD_EXCEL_ERROR_COUNT, 1L);
                            getErrorResultHolder().put(currentLineCount, result);
                        }
                    }
                };
        //放入线程池
        getThreadPool().submit(thread);
    }*/
}
