package com.ooooor.xls4j.infrastructure.util;

import org.xml.sax.helpers.DefaultHandler;

/**
 * @description:
 * @author: chenr
 * @date: 18-11-9
 */
public class XlsxSheetHandler extends DefaultHandler implements RowDataHandler {

    /*private final static int ASCII_A_INT = 'A';
    private final static int ASCII_0_INT = '0';
    private final static int ASCII_9_INT = '9';

    private final SharedStringsTable sst;
    private String lastContents;
    private boolean nextIsString;
    private boolean inlineStr;
    private int colIndex = 0;
    private final LruCache<Integer, String> lruCache = new LruCache<>(50);
    private Object[] currentRowData = new Object[MAX_COL_NUM];
    private boolean onlyCountLine = false;

    // 总行数
    @Getter
    private long totalRowCount;

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
            totalRowCount++;
            if (!onlyCountLine) {
                handleRowData();
                //清空行数据
                Arrays.fill(currentRowData, 0, currentRowData.length, null);
            }

        } else if ("worksheet".equals(name)) {// Sheet读取完成
            //
            System.out.println("total lines: " + totalRowCount);
        }
    }

    @Override
    public void characters(char[] ch, int start, int length) throws SAXException { // NOSONAR
        lastContents += new String(ch, start, length);
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
        return headerIndexMap;
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
