package com.ooooor.xls4j.infrastructure.util;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.google.common.primitives.Longs;
import com.ooooor.xls4j.application.dto.ImportResultDto;
import com.ooooor.xls4j.application.dto.OneLineResultDto;
import com.ooooor.xls4j.application.service.DetailImportService;
import com.ooooor.xls4j.domain.exception.ServiceException;
import com.ooooor.xls4j.infrastructure.controller.acceptor.FileAcceptorController;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.eventusermodel.FormatTrackingHSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.eventusermodel.MissingRecordAwareHSSFListener;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.eventusermodel.dummyrecord.MissingCellDummyRecord;
import org.apache.poi.hssf.record.BlankRecord;
import org.apache.poi.hssf.record.BoolErrRecord;
import org.apache.poi.hssf.record.BoundSheetRecord;
import org.apache.poi.hssf.record.FormulaRecord;
import org.apache.poi.hssf.record.LabelRecord;
import org.apache.poi.hssf.record.LabelSSTRecord;
import org.apache.poi.hssf.record.NoteRecord;
import org.apache.poi.hssf.record.NumberRecord;
import org.apache.poi.hssf.record.RKRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.SSTRecord;
import org.apache.poi.hssf.record.StringRecord;
import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.springframework.data.redis.core.RedisTemplate;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

import javax.xml.parsers.ParserConfigurationException;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;

/**
 *
 */
public class ExcelImportService {

    public final static int MAX_COL_NUM = 64;

    /**
     * 表头: Map<列名,列索引>
     */
    private Map<String, Integer> headerIndexMap;
    /**
     * 导入失败返回的信息
     */
    private List<OneLineResultDto> errorLines;

    private final static int ASCII_A_INT = 'A';
    private final static int ASCII_0_INT = '0';
    private final static int ASCII_9_INT = '9';
    /**
     * 总行数
     */
    private long lineCount = 0;

    private String importTrxId;

    private DetailImportService detailImportService;

    private RedisTemplate<String, String> redisTemplate;

    private Map<Long, OneLineResultDto> errorResultHolder = new ConcurrentHashMap<>(16384);

    private ExecutorService threadPool;

    private boolean isMultipleThread = false;

    public ExcelImportService(Class detailImportServiceClass, String importTrxId, RedisTemplate<String, String> redisTemplate) {
        this.detailImportService = (DetailImportService) SpringBeanUtil.getBean(detailImportServiceClass);
        this.importTrxId = importTrxId;
        this.redisTemplate = redisTemplate;
    }

    public ExcelImportService(Class detailImportServiceClass, String importTrxId, RedisTemplate<String, String> redisTemplate, boolean isMultipleThread) {
        this(detailImportServiceClass, importTrxId, redisTemplate);
        this.isMultipleThread = isMultipleThread;
    }

    /**
     * 处理xlsx格式excel
     *
     * @param filename
     * @return
     * @throws Exception
     */
    public ImportResultDto processXlsx(String filename) throws Exception {
        init();
        try (OPCPackage pkg = OPCPackage.open(filename, PackageAccess.READ)) {
            XSSFReader r = new XSSFReader(pkg);
            SharedStringsTable sst = r.getSharedStringsTable();

            XMLReader parser = fetchSheetParser(sst);

            // process the first sheet
            try (InputStream sheet = r.getSheetsData().next()) {
                InputSource sheetSource = new InputSource(sheet);
                parser.parse(sheetSource);
                while (((ThreadPoolExecutor) this.threadPool).getActiveCount() > 0) {
                    Thread.sleep(2 * 1000);
                }
                Object successCount = redisTemplate.opsForHash().get(importTrxId, FileAcceptorController.HASH_FIELD_EXCEL_SUCCESS_COUNT);
                //扣减excel表头
                Long importTotalCount = Longs.tryParse(String.valueOf(successCount)) - 1;
                redisTemplate.opsForHash().put(importTrxId, FileAcceptorController.HASH_FIELD_IMPORT_TOTAL_COUNT, String.valueOf(importTotalCount));
                threadPool.shutdown();
                Set<OneLineResultDto> errorLineSet = new TreeSet<OneLineResultDto>(Comparator.comparing(OneLineResultDto::getLineNum));

                errorResultHolder.forEach((k, v) ->
                    errorLineSet.add(v)
                );
                errorLines = Lists.newArrayList(errorLineSet);
                return new ImportResultDto(headerIndexMap, errorLines, lineCount - errorLines.size() - 1);
            }
        }
    }

    /**
     * 获取xlsx格式excel的总行数
     *
     * @param filename
     * @return
     * @throws Exception
     */
    public long getTotalCountXlsx(String filename) throws Exception {
        init();
        try (OPCPackage pkg = OPCPackage.open(filename, PackageAccess.READ)) {
            XSSFReader r = new XSSFReader(pkg);
            SharedStringsTable sst = r.getSharedStringsTable();

            XMLReader parser = fetchSheetParser(sst, true);

            // process the first sheet
            try (InputStream sheet = r.getSheetsData().next()) {
                InputSource sheetSource = new InputSource(sheet);
                parser.parse(sheetSource);

                return this.getLineCount();
            }
        }
    }

    private void init() {
        headerIndexMap = Maps.newHashMap();
        errorLines = Lists.newArrayList();
        lineCount = 0;
        if (threadPool == null) {
            int cpuNum = Runtime.getRuntime().availableProcessors();
            if (this.isMultipleThread) {
                threadPool = new ThreadPoolExecutor(cpuNum + 1, cpuNum * 2,
                        60L, TimeUnit.SECONDS, new ArrayBlockingQueue<>(10000), Executors.defaultThreadFactory(), new ThreadPoolExecutor.CallerRunsPolicy());
            } else {
                threadPool = new ThreadPoolExecutor(1, 1,
                        60L, TimeUnit.SECONDS, new ArrayBlockingQueue<>(10000), Executors.defaultThreadFactory(), new ThreadPoolExecutor.CallerRunsPolicy());
            }

        }
    }

    /**
     * 处理xls格式excel
     *
     * @param fileName
     * @return
     * @throws Exception
     */
    public ImportResultDto processXls(String fileName) throws Exception {
        init();
        XlsSheetHandler handler = new XlsSheetHandler();
        MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(handler);
        FormatTrackingHSSFListener formatListener = new FormatTrackingHSSFListener(listener);
        handler.setFormatListener(formatListener);
        HSSFEventFactory factory = new HSSFEventFactory();
        HSSFRequest request = new HSSFRequest();
        request.addListenerForAllRecords(formatListener);
        factory.processWorkbookEvents(request, new POIFSFileSystem(new FileInputStream(fileName)));

        while (((ThreadPoolExecutor) this.threadPool).getActiveCount() > 0) {
            Thread.sleep(2 * 1000);
        }

        Object successCount = redisTemplate.opsForHash().get(importTrxId, FileAcceptorController.HASH_FIELD_EXCEL_SUCCESS_COUNT);
        Long importTotalCount = Longs.tryParse(String.valueOf(successCount)) - 1;
        redisTemplate.opsForHash().put(importTrxId, FileAcceptorController.HASH_FIELD_IMPORT_TOTAL_COUNT, String.valueOf(importTotalCount));
        threadPool.shutdown();
        Set<OneLineResultDto> errorLineSet = new TreeSet<OneLineResultDto>(Comparator.comparing(OneLineResultDto::getLineNum));

        errorResultHolder.forEach((k, v) ->
            errorLineSet.add(v)
        );
        errorLines = Lists.newArrayList(errorLineSet);
        return new ImportResultDto(headerIndexMap, errorLines, lineCount - errorLines.size() - 1);
    }

    /**
     * 处理xls格式excel
     *
     * @param fileName
     * @return
     * @throws Exception
     */
    public long getTotalCountlsx(String fileName) throws Exception {
        lineCount = 0;
        XlsSheetHandler handler = new XlsSheetHandler(true);
        MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(handler);
        FormatTrackingHSSFListener formatListener = new FormatTrackingHSSFListener(listener);
        handler.setFormatListener(formatListener);
        HSSFEventFactory factory = new HSSFEventFactory();
        HSSFRequest request = new HSSFRequest();
        request.addListenerForAllRecords(formatListener);
        factory.processWorkbookEvents(request, new POIFSFileSystem(new FileInputStream(fileName)));

        return getLineCount();
    }

    /**
     * 获取操作类型
     * @return
     */
    public Byte getOperateType() {
        return detailImportService.getOperateType();
    }


    private XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException, ParserConfigurationException {
        XMLReader parser = SAXHelper.newXMLReader();
        ContentHandler handler = new XlsxSheetHandler(sst);
        parser.setContentHandler(handler);
        return parser;
    }

    private XMLReader fetchSheetParser(SharedStringsTable sst, boolean onlyCountLine) throws SAXException, ParserConfigurationException {
        XMLReader parser = SAXHelper.newXMLReader();
        ContentHandler handler = new XlsxSheetHandler(sst, onlyCountLine);
        parser.setContentHandler(handler);
        return parser;
    }

    private interface RowDataHandler {
        long getLineCount();

        Object[] getCurrentRowData();

        DetailImportService getDetailImportService();

        Map<String, Integer> getHeaderIndexMap();

        Map<Long, OneLineResultDto> getErrorResultHolder();

        RedisTemplate<String, String> getRedisTemplate();

        String getImportTrxId();

        ExecutorService getThreadPool();

        /*()*
         * 处理行数据
         */
        default void handleRowData() {
            if (getLineCount() == 1) {
                for (int i = 0; i < getCurrentRowData().length; i++) {
                    if (getCurrentRowData()[i] == null) {
                        break;
                    }

                    getHeaderIndexMap().put(getCurrentRowData()[i].toString(), i);
                }
                getRedisTemplate().opsForHash().increment(getImportTrxId(), FileAcceptorController.HASH_FIELD_EXCEL_SUCCESS_COUNT, 1L);
                return;
            }
            long currentLineCount = getLineCount();
            List<Object> list = Lists.newArrayList(Arrays.copyOf(getCurrentRowData(), getHeaderIndexMap().size()));

            Thread thread =
                    new Thread(() -> {
                        //处理一行记录
                        OneLineResultDto result = getDetailImportService().doImport(getHeaderIndexMap(), list, currentLineCount, getImportTrxId());
                        if (null != result && StringUtils.isNotBlank(result.getResult())) {
                            getRedisTemplate().opsForHash().increment(getImportTrxId(), FileAcceptorController.HASH_FIELD_EXCEL_SUCCESS_COUNT, 1L);
                            getErrorResultHolder().put(currentLineCount, result);
                        } else {
                            getRedisTemplate().opsForHash().increment(getImportTrxId(), FileAcceptorController.HASH_FIELD_EXCEL_ERROR_COUNT, 1L);
                            // getErrorResultHolder().put(currentLineCount, result);
                        }
                    });
            //放入线程池
            getThreadPool().submit(thread);
        }
    }

    /**
     * 处理xlsx格式excel表格
     */
    private class XlsxSheetHandler extends DefaultHandler implements RowDataHandler {
        private final SharedStringsTable sst;
        private String lastContents;
        private boolean nextIsString;
        private boolean inlineStr;
        private int colIndex = 0;
        private final LruCache<Integer, String> lruCache = new LruCache<>(50);
        private Object[] currentRowData = new Object[MAX_COL_NUM];
        private boolean onlyCountLine = false;

        private class LruCache<A, B> extends LinkedHashMap<A, B> {
            private final int maxEntries;

            public LruCache(final int maxEntries) {
                super(maxEntries + 1, 1.0f, true);
                this.maxEntries = maxEntries;
            }

            @Override
            protected boolean removeEldestEntry(final Map.Entry<A, B> eldest) {
                return super.size() > maxEntries;
            }
        }

        private XlsxSheetHandler(SharedStringsTable sst) {
            this.sst = sst;
        }

        private XlsxSheetHandler(SharedStringsTable sst, boolean onlyCountLine) {
            this.sst = sst;
            this.onlyCountLine = onlyCountLine;
        }

        @Override
        public void startElement(String uri, String localName, String name,
                                 Attributes attributes) throws SAXException {
            if ("c".equals(name)) {// <row>:开始处理某一行
                String cellType = attributes.getValue("t");
                nextIsString = cellType != null && cellType.equals("s");
                inlineStr = cellType != null && cellType.equals("inlineStr");
                String cellDateType = attributes.getValue("s");
                colIndex = getCurrentColIndex(attributes);
            } else if ("row".equals(name)) {// <c>:一个单元格
                //
            } else if ("v".equals(name)) {// <v>:单元格值
                //
            } else if ("f".equals(name)) {// <f>:公式表达式标签
                //
            } else if ("is".equals(name)) {// 内联字符串外部标签
                //
            } else if ("col".equals(name)) {// 处理隐藏列c
                //
            }

            // Clear contents cache
            lastContents = "";
        }

        private int getCurrentColIndex(Attributes attributes) {
            String val = attributes.getValue("r");
            int len = val.length();
            int numIndex = 0;

            int index = 1;
            for (int i = 0; i < len; i++) {
                if (ASCII_0_INT <= val.charAt(i) && val.charAt(i) <= ASCII_9_INT) {
                    numIndex = i;
                    break;
                }
            }
            val = val.substring(0, numIndex);
            if (val.length() == 1) {
                return val.charAt(0) - ASCII_A_INT;
            } else if (val.length() == 2) {
                index = (val.charAt(0) - ASCII_A_INT + 1) * 26 + val.charAt(1) - ASCII_A_INT;
                if (index >= MAX_COL_NUM) {
                    throw new ServiceException("当前列" + val + ",超过了允许的最大列数,请检查表格是否符合模板要求");
                }
                return index;
            } else {
                throw new ServiceException("当前列" + val + ",超过了允许的最大列数,请检查表格是否符合模板要求");
            }
        }

        @Override
        public void endElement(String uri, String localName, String name)
                throws SAXException {
            if (nextIsString) {
                Integer idx = Integer.valueOf(lastContents);
                lastContents = lruCache.get(idx);
                if (lastContents == null && !lruCache.containsKey(idx)) {
                    lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
                    lruCache.put(idx, lastContents);
                }
                nextIsString = false;
            }

            if (name.equals("v") || (inlineStr && name.equals("c"))) {
                currentRowData[colIndex] = StringUtils.trim(lastContents);
            } else if (name.equals("row")) {
                lineCount++;
                if (!onlyCountLine) {
                    handleRowData();
                    //清空行数据
                    Arrays.fill(currentRowData, 0, currentRowData.length, null);
                }

            } else if ("worksheet".equals(name)) {// Sheet读取完成
                //
                System.out.println("total lines: " + lineCount);
            }
        }

        @Override
        public void characters(char[] ch, int start, int length) throws SAXException { // NOSONAR
            lastContents += new String(ch, start, length);
        }

        @Override
        public long getLineCount() {
            return lineCount;
        }


        @Override
        public Object[] getCurrentRowData() {
            return currentRowData;
        }

        @Override
        public DetailImportService getDetailImportService() {
            return detailImportService;
        }

        @Override
        public Map<String, Integer> getHeaderIndexMap() {
            return headerIndexMap;
        }

        @Override
        public Map<Long, OneLineResultDto> getErrorResultHolder() {
            return errorResultHolder;
        }

        @Override
        public String getImportTrxId() {
            return importTrxId;
        }

        @Override
        public ExecutorService getThreadPool() {
            return threadPool;
        }

        @Override
        public RedisTemplate<String, String> getRedisTemplate() {
            return redisTemplate;
        }
    }

    private class XlsSheetHandler implements HSSFListener, RowDataHandler {

        private ArrayList<Record> boundSheetRecords = Lists.newArrayList();
        private int nextColumn;
        private SSTRecord sstRecord;
        //当前需要的表单index
        private boolean outputNextStringRecord;
        private FormatTrackingHSSFListener formatListener;
        //当前行数据
        private Object[] currentRowData = new Object[MAX_COL_NUM];
        private boolean onlyCountLine = false;

        public XlsSheetHandler(boolean onlyCountLine) {
            this.onlyCountLine = onlyCountLine;
        }

        public XlsSheetHandler() {
        }

        @Override
        public long getLineCount() {
            return lineCount;
        }

        @Override
        public Object[] getCurrentRowData() {
            return currentRowData;
        }

        @Override
        public DetailImportService getDetailImportService() {
            return detailImportService;
        }

        @Override
        public Map<String, Integer> getHeaderIndexMap() {
            return headerIndexMap;
        }

        @Override
        public Map<Long, OneLineResultDto> getErrorResultHolder() {
            return errorResultHolder;
        }

        @Override
        public String getImportTrxId() {
            return importTrxId;
        }

        @Override
        public ExecutorService getThreadPool() {
            return threadPool;
        }

        @Override
        public RedisTemplate<String, String> getRedisTemplate() {
            return redisTemplate;
        }

        public void setFormatListener(FormatTrackingHSSFListener formatListener) {
            this.formatListener = formatListener;
        }

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
                        thisColumn = nextColumn;
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
                lineCount++;
                if (!onlyCountLine) {
                    handleRowData();
                    //清空行数据
                    Arrays.fill(currentRowData, 0, currentRowData.length, null);
                }
            }
        }
    }

    public long getLineCount() {
        return lineCount;
    }
}
