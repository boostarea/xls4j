package com.ooooor.xls4j.infrastructure.util;

/**
 * @description:
 * @author: chenr
 * @date: 18-11-9
 */
// @NoArgsConstructor
// @RequiredArgsConstructor
public class XlsSheetHandler implements RowDataHandler {

   /* public final static int MAX_COL_NUM = 64;

    // 总行数
    @Getter
    private long totalRowCount;

    private List<Record> boundSheetRecords = Lists.newArrayList();

    private SSTRecord sstRecord;

    //当前需要的表单index
    private boolean outputNextStringRecord;
    private FormatTrackingHSSFListener formatListener;
    //当前行数据
    private Object[] currentRowData = new Object[MAX_COL_NUM];

    private boolean onlyCountLine = false;

    @Override
    public void processRecord(Record record) {
        int thisColumn = -1;
        String value = null;

        switch (record.getSid()) {
            case BoundSheetRecord.sid:
                boundSheetRecords.add(record);
                break;
            case SSTRecord.sid:
                sstRecord = (SSTRecord) record;
                break;
            case BlankRecord.sid:
                BlankRecord brec = (BlankRecord) record;

                thisColumn = brec.getColumn();
                break;
            case BoolErrRecord.sid:
                BoolErrRecord berec = (BoolErrRecord) record;

                thisColumn = berec.getColumn();
                break;

            case FormulaRecord.sid:
                FormulaRecord frec = (FormulaRecord) record;

                thisColumn = frec.getColumn();

                break;
            case StringRecord.sid:
                if (outputNextStringRecord) {
                    outputNextStringRecord = false;
                }
                break;
            case LabelRecord.sid:
                LabelRecord lrec = (LabelRecord) record;

                thisColumn = lrec.getColumn();
                value = lrec.getValue().trim();
                currentRowData[thisColumn] = StringUtils.trim(value);
                break;
            case LabelSSTRecord.sid:
                LabelSSTRecord lsrec = (LabelSSTRecord) record;

                thisColumn = lsrec.getColumn();
                if (sstRecord == null) {
                    currentRowData[thisColumn] = StringUtils.trim(value);
                } else {
                    value = sstRecord.getString(lsrec.getSSTIndex()).toString().trim();
                    currentRowData[thisColumn] = StringUtils.trim(value);
                }
                break;
            case NoteRecord.sid:
                NoteRecord nrec = (NoteRecord) record;
                thisColumn = nrec.getColumn();
                break;
            case NumberRecord.sid:
                NumberRecord numrec = (NumberRecord) record;

                thisColumn = numrec.getColumn();
                value = formatListener.formatNumberDateCell(numrec).trim();
                currentRowData[thisColumn] = StringUtils.trim(value);
                break;
            case RKRecord.sid:
                RKRecord rkrec = (RKRecord) record;

                thisColumn = rkrec.getColumn();
                break;
            default:
                break;
        }
        if (thisColumn > -1 && value == null) {
            currentRowData[thisColumn] = null;
        }
        // 空值的操作
        if (record instanceof MissingCellDummyRecord) {
            MissingCellDummyRecord mc = (MissingCellDummyRecord) record;
            thisColumn = mc.getColumn();
            currentRowData[thisColumn] = null;
        }
        // 行结束时的操作
        if (record instanceof LastCellOfRowDummyRecord) {
            totalRowCount++;
            if (!onlyCountLine) {
                handleRowData();
                //清空行数据
                Arrays.fill(currentRowData, 0, currentRowData.length, null);
            }
        }
    }

    @Override
    public long getLineCount() {
        return totalRowCount;
    }

    @Override
    public Object[] getCurrentRowData() {
        return new Object[0];
    }

    @Override
    public DetailImportService getDetailImportService() {
        return null;
    }

    @Override
    public Map<String, Integer> getHeaderIndexMap() {
        return null;
    }

    @Override
    public Map<Long, OneLineResultDto> getErrorResultHolder() {
        return null;
    }

    @Override
    public RedisTemplate<String, String> getRedisTemplate() {
        return null;
    }

    @Override
    public String getImportTrxId() {
        return null;
    }

    @Override
    public ExecutorService getThreadPool() {
        return null;
    }
*/}
