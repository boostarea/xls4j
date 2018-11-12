package com.ooooor.xls4j.infrastructure.controller.acceptor;

import com.google.common.base.Splitter;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.google.common.primitives.Longs;
import com.ooooor.xls4j.application.dto.AjaxRes;
import com.ooooor.xls4j.application.dto.ImportResultDto;
import com.ooooor.xls4j.application.dto.OneLineResultDto;
import com.ooooor.xls4j.application.service.OutOrderImportServiceImpl;
import com.ooooor.xls4j.domain.exception.ServiceException;
import com.ooooor.xls4j.infrastructure.util.ExcelImportService;
import com.ooooor.xls4j.infrastructure.util.UuidUtil;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.time.StopWatch;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.data.redis.core.RedisTemplate;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Comparator;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.TreeSet;

/**
 * @description:
 * @author: chenr
 * @date: 18-11-9
 */
@RestController
@RequestMapping("acceptor")
public class FileAcceptorController {

    private Logger logger = LoggerFactory.getLogger(FileAcceptorController.class);

    private static final String XLS = "xls";

    private static final String XLSX = "xlsx";

    protected final Map<String, Class> detailImportServiceMap = Maps.newConcurrentMap();
    /**
     * 导入源文件名
     */
    public static final String HASH_FIELD_EXCEL_SOURCE_FILE ="sourceFile";
    /**
     * EXCEL导入总行数（包含表头）
     */
    public static final String HASH_FIELD_EXCEL_TOTAL_COUNT ="excelTotalCount";
    /**
     * EXCEL导入校验错误的行数
     */
    public static final String HASH_FIELD_EXCEL_ERROR_COUNT ="excelErrorCount";
    /**
     * EXCEL导入校验通过的行数
     */
    public static final String HASH_FIELD_EXCEL_SUCCESS_COUNT ="excelSuccessCount";
    /**
     * 请求参数
     */
    public static final String REQUEST_PARAMETER_KEY ="requestParameter";

    public static final String HASH_FIELD_IMPORT_TOTAL_COUNT ="importTotalCount";
    public static final String HASH_FIELD_IMPORT_ERROR_COUNT ="importErrorCount";
    public static final String HASH_FIELD_IMPORT_SUCCESS_COUNT ="importSuccessCount";
    public static final String KEY_IMPORT_ERROR_ROWS ="%s_importErrorRows";
    public static final String HASH_FIELD_RESULT_FILE ="resultFile";

    public static final String HASH_KEY_IMPORT_PROGRESS_PREFIX ="IMPORTTRX_ID_";


    @Qualifier("LineResultRedisTemplate")
    @Autowired
    private RedisTemplate<String, OneLineResultDto> redisTemplateForErrorRows;

    @Autowired
    @Qualifier("LineMapRedisTemplate")
    private RedisTemplate<String, Map<String, String[]>> requestParameterRedisTemplate;

    @Autowired
    protected RedisTemplate<String, String> redisTemplate;

    @RequestMapping("upload")
    public AjaxRes upload(@RequestParam("file")MultipartFile file, HttpServletRequest request, String importType) throws UnsupportedEncodingException {
        request.setCharacterEncoding("utf-8");
        AjaxRes result = new AjaxRes();
        File tmpFile = null;
        try {
            //获取文件后缀名
            String suffix = file.getOriginalFilename().substring(file.getOriginalFilename().lastIndexOf(".") + 1);
            if (!"XLS".contains(suffix.toLowerCase()) && !XLSX.contains(suffix.toLowerCase())) {
                result.setMsg("请上传excel表格(xls/xlsx格式)");
                return result;
            }

            String realPath = System.getProperty("java.io.tmpdir");
            String fileName = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMddHHmmss")) + UuidUtil.get32UUID().hashCode();
            File baseFile = new File(realPath);
            tmpFile = new File(baseFile, fileName + "." + suffix);
            if (!baseFile.exists()) {
                baseFile.mkdirs();
            }
            //保存临时文件
            file.transferTo(tmpFile);
            // Class importService = detailImportServiceMap.get(importType);
            Class importService = OutOrderImportServiceImpl.class;
            Optional.ofNullable(importService).orElseThrow(() -> new ServiceException("can not found importHandler: " + importType));
            String importTrxId = HASH_KEY_IMPORT_PROGRESS_PREFIX + fileName;
            ExcelImportService excelImportService = new ExcelImportService(importService, importTrxId, redisTemplate);

            Long totalCount;
            if (XLS.contains(suffix.toLowerCase())) {
                totalCount = excelImportService.getTotalCountlsx(tmpFile.getAbsolutePath());
            } else {
                totalCount = excelImportService.getTotalCountXlsx(tmpFile.getAbsolutePath());
            }
            redisTemplate.opsForHash().put(importTrxId, HASH_FIELD_EXCEL_SOURCE_FILE, tmpFile.getName());
            redisTemplate.opsForHash().put(importTrxId, HASH_FIELD_EXCEL_TOTAL_COUNT, String.valueOf(totalCount));
            redisTemplate.opsForHash().put(importTrxId, HASH_FIELD_EXCEL_ERROR_COUNT, "0");
            redisTemplate.opsForHash().put(importTrxId, HASH_FIELD_EXCEL_SUCCESS_COUNT, "0");

            //请求参数
            if(null != request.getParameterMap() && request.getParameterMap().size() > 0) {
                requestParameterRedisTemplate.opsForValue().set(REQUEST_PARAMETER_KEY + importTrxId, request.getParameterMap());
            }
            //文件保存成功后，返回
            result.setSucceed(importTrxId);

            new ExcelParseThread(importType, tmpFile, importTrxId).start();

            return result;

        } catch (Exception e) {
            logger.error("导入excel出错:" + e.getMessage(), e);
            result.setMsg("导入excel出错: " + e.getMessage());
            return result;
        }
    }


    class ExcelParseThread extends Thread {

        private String importType;

        private File tmpFile;

        private String importTrxId;

        public ExcelParseThread(String importType, File tmpFile, String importTrxId) {
            this.importType = importType;
            this.tmpFile = tmpFile;
            this.importTrxId = importTrxId;
        }

        @Override
        public void run() {
            StopWatch stopWatch = new StopWatch();
            stopWatch.start();
            ImportResultDto resultDto;

            Class importService = OutOrderImportServiceImpl.class;
            Optional.ofNullable(importService).orElseThrow(() -> new ServiceException("can not found importHandler: " + importType));

            ExcelImportService excelImportService = new ExcelImportService(importService, importTrxId, redisTemplate, true);
            String suffix = tmpFile.getName().substring(tmpFile.getName().lastIndexOf(".") + 1);

            try {
                if (XLS.contains(suffix.toLowerCase())) {
                    resultDto = excelImportService.processXls(tmpFile.getAbsolutePath());
                } else {
                    resultDto = excelImportService.processXlsx(tmpFile.getAbsolutePath());
                }
                stopWatch.stop();
                logger.info("解析验证excel超作耗时：" + stopWatch.getTime());
                while(true) {

                    List<Object> allCounts =
                            redisTemplate.opsForHash().multiGet(importTrxId, Lists.newArrayList(HASH_FIELD_IMPORT_SUCCESS_COUNT,
                                    HASH_FIELD_IMPORT_ERROR_COUNT, HASH_FIELD_IMPORT_TOTAL_COUNT));

                    long successCount = Longs.tryParse(Objects.toString(allCounts.get(0), "0"));
                    long errorCount = Longs.tryParse(Objects.toString(allCounts.get(1), "0"));
                    long totalCount = Longs.tryParse(Objects.toString(allCounts.get(2), "0"));

                    if(successCount + errorCount >= totalCount) {
                        break;
                    }
                    Thread.sleep(2*1000);
                }
                String errorRowsKey = String.format(KEY_IMPORT_ERROR_ROWS, importTrxId);

                List<OneLineResultDto> lineResultDtoList = redisTemplateForErrorRows.opsForList().range(errorRowsKey, 0, -1);

                if(CollectionUtils.isNotEmpty(lineResultDtoList)) {
                    Set<OneLineResultDto> errorLineSet = new TreeSet<OneLineResultDto>(Comparator.comparing(OneLineResultDto::getLineNum));

                    lineResultDtoList.forEach(errorLine ->
                            errorLineSet.add(errorLine)
                    );
                    if(CollectionUtils.isNotEmpty(resultDto.getErrorLines())) {
                        resultDto.getErrorLines().forEach(errorLine ->
                                errorLineSet.add(errorLine)
                        );
                    }
                    resultDto.setErrorLines(Lists.newArrayList(errorLineSet));
                }

                //导入结果excel
                String fileName = Splitter.on(HASH_KEY_IMPORT_PROGRESS_PREFIX).splitToList(importTrxId).get(1);
                File resultFile = new File(tmpFile.getParentFile(), "importResult_" + fileName + "ct" + stopWatch.getTime() + "." + suffix);
                stopWatch.reset();
                stopWatch.start();
                try (FileOutputStream fileOutputStream = new FileOutputStream(resultFile)) {
                    writeToFile(resultDto, fileOutputStream);
                    redisTemplate.opsForHash().put(importTrxId, HASH_FIELD_RESULT_FILE, resultFile.getName());
                }
                stopWatch.stop();
                logger.info("生成导入结果excel超作耗时：" + stopWatch.getTime());
                redisTemplateForErrorRows.delete(errorRowsKey);
                //请求参数
                requestParameterRedisTemplate.delete(REQUEST_PARAMETER_KEY + importTrxId);

            }catch(Exception e) {
                logger.error(e.getMessage(), e);
            }
        }
    }

    @RequestMapping("downloadImportResult")
    public void downloadImportResult(String fileName, HttpServletResponse response) throws IOException {
        File resultFile = new File(System.getProperty("java.io.tmpdir") + File.separator + fileName);

        writeToResponse(response, resultFile);
    }

    private void writeToResponse(HttpServletResponse response, File file) throws IOException {
        try (BufferedInputStream bis = new BufferedInputStream(new FileInputStream(file));
             BufferedOutputStream bos = new BufferedOutputStream(response.getOutputStream())) {
            // 设置response参数，可以打开下载页面
            response.reset();
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8");
            response.setHeader("Content-Disposition", "attachment;filename="+ new String(file.getName().getBytes(), "UTF-8"));

            byte[] buff = new byte[8192];
            int bytesRead;
            while (-1 != (bytesRead = bis.read(buff, 0, buff.length))) {
                bos.write(buff, 0, bytesRead);
            }
            bos.flush();
        }
    }

    protected void writeToFile(ImportResultDto resultDto, OutputStream out) throws IOException {
        try (SXSSFWorkbook workbook = new SXSSFWorkbook(resultDto.getErrorLines().size() + 1)) {
            // 声明一个工作薄
            workbook.setCompressTempFiles(true);
            // 列头样式
            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.setAlignment(HorizontalAlignment.CENTER);
            Font headerFont = workbook.createFont();
            headerFont.setFontHeightInPoints((short) 12);
            headerFont.setBold(true);
            headerStyle.setFont(headerFont);

            // 生成一个(带标题)表格
            SXSSFSheet sheet = workbook.createSheet("导入结果");

            if (CollectionUtils.isEmpty(resultDto.getErrorLines())) {
                sheet.createRow(0).createCell(0).setCellValue("成功导入"+resultDto.getSuccessCount()+"条");
            } else {
                //表头行
                SXSSFRow headerRow = sheet.createRow(0);
                Map<String, Integer> header = resultDto.getHeaderIndexMap();
                header.keySet().forEach(colName ->  {
                    SXSSFCell cell = headerRow.createCell(header.get(colName) + 1);
                    cell.setCellStyle(headerStyle);
                    cell.setCellValue(colName);
                });
                SXSSFCell headerRowCell = headerRow.createCell(0);
                headerRowCell.setCellStyle(headerStyle);
                headerRowCell.setCellValue("所在行数");
                headerRowCell = headerRow.createCell(header.size() + 1);
                headerRowCell.setCellStyle(headerStyle);
                headerRowCell.setCellValue("异常信息");

                // 单元格样式
                CellStyle cellStyle = workbook.createCellStyle();
                cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

                //异常列格式
                CellStyle errCellStyle = workbook.createCellStyle();
                Font errCellFont = workbook.createFont();
                errCellFont.setColor(Font.COLOR_RED);
                errCellStyle.setFont(errCellFont);

                //异常记录行
                for (int i = 1; i < resultDto.getErrorLines().size() + 1; i++) {
                    SXSSFRow dataRow = sheet.createRow(i);
                    OneLineResultDto line = resultDto.getErrorLines().get(i - 1);

                    //所在行数
                    SXSSFCell firstCell = dataRow.createCell(0);
                    firstCell.setCellStyle(cellStyle);
                    firstCell.setCellValue(line.getLineNum());

                    //商品信息
                    header.keySet().forEach(colName -> {
                        SXSSFCell dataRowCell = dataRow.createCell(header.get(colName) + 1);
                        dataRowCell.setCellStyle(cellStyle);
                        dataRowCell.setCellValue(Optional.ofNullable(line.getLineData().get(header.get(colName)))
                                .orElse("").toString());
                    });

                    //异常信息
                    SXSSFCell errCell = dataRow.createCell(header.size() + 1);
                    errCell.setCellStyle(errCellStyle);
                    errCell.setCellValue(line.getErrMsg());
                }
            }

            workbook.write(out);
        }
    }
}
