package com.ooooor.exl4j.util;

import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.ooooor.exl4j.model.JoomOrder;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.Writer;

/**
 * 通过解析Exl，拼接生成sql语句
 */
public class ExlReader {

    private static final String STR_FORMAT = "update seller_order t1, seller_order_detail t2 set t2.out_extra_price=%s where t1.id=t2.order_id and t1.platform_order_no='%s';";
    // 待解析文件所在路径
    private static final String IN_PATH = "/Users/chen/Desktop/joom/";
    // 生成sql文件
    private static final String OUT_FILE = "joom.sql";

    public static void main(String[] args) throws IOException {
        long startTime = System.currentTimeMillis();
        InputStream[] inputStreams = getInputStream(IN_PATH);
        if (null == inputStreams) {
            System.out.println("Error,not generated input stream");
        }

        File file = createFile(OUT_FILE);
        try (Writer out = new FileWriter(file)) {
            final StringBuffer buffer = new StringBuffer();
            for (InputStream inputStream : inputStreams) {
                try {
                    ExcelReader reader = new ExcelReader(inputStream, ExcelTypeEnum.XLS, null,
                            new AnalysisEventListener<JoomOrder>() {
                                @Override
                                public void invoke(JoomOrder object, AnalysisContext context) {
                                    buffer.append(String.format(STR_FORMAT, object.getExtraPrice(), object.getOrderId()));
                                    buffer.append(System.getProperty("line.separator"));
                                }

                                @Override
                                public void doAfterAllAnalysed(AnalysisContext context) {
                                }
                            });

                    reader.read(new Sheet(1, 1, JoomOrder.class));
                } catch (Exception e) {
                    e.printStackTrace();

                } finally {
                    try {
                        inputStream.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            }
            out.write(buffer.toString());
        } catch (Exception e) {
            // TODO: handle exception
            e.printStackTrace();
        }
        System.out.println("Done, time:" + (System.currentTimeMillis() - startTime));
    }

    private static InputStream[] getInputStream(String fileName) throws FileNotFoundException {
        File[] files = new File(fileName).listFiles();
        if (null == files) {
            return null;
        }

        InputStream[] streams = new InputStream[files.length];
        for (int i = 0; i < files.length; i++) {
            streams[i] = new FileInputStream(files[i]);
        }
        return streams;
        // return Thread.currentThread().getContextClassLoader().getResourceAsStream("" + fileName);
    }

    private static File createFile(String fileName) {
        File file = new File(fileName);
        if (!file.exists()) {
            try {
                file.createNewFile();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return file;
    }

}
