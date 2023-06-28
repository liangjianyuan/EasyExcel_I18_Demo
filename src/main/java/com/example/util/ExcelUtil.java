package com.example.util;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.data.ReadCellData;
import com.alibaba.excel.metadata.property.ExcelContentProperty;
import com.alibaba.excel.read.metadata.property.ExcelReadHeadProperty;
import com.alibaba.excel.util.ConverterUtils;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.handler.WriteHandler;
import com.alibaba.excel.write.style.column.LongestMatchColumnWidthStyleStrategy;
import com.example.config.SpringBeanUtil;
import com.example.handler.SimpleCellWriteHandler;
import com.example.handler.SimpleSheetWriteHandler;
import org.apache.commons.collections4.MapUtils;
import org.apache.commons.lang3.StringUtils;
import org.springframework.context.MessageSource;
import org.springframework.context.i18n.LocaleContextHolder;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.*;

public class ExcelUtil {

    /**
     * 导出初始化response
     * @param response
     * @throws UnsupportedEncodingException
     */
    private static void initResponse(HttpServletResponse response,String fileName,boolean flag) throws UnsupportedEncodingException {
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setCharacterEncoding("utf-8");
        //国际化
        if(flag){
            fileName = SpringBeanUtil.getBean(MessageSource.class).getMessage(fileName,null, LocaleContextHolder.getLocale());
        }
        fileName = URLEncoder.encode(fileName, "UTF-8").replaceAll("\\+", "%20");
        response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
    }

    /**
     * 构建Web下导出的Builder
     * @param response
     * @param fileName
     * @param flag
     * @param writeHandlerList
     * @return
     * @throws IOException
     */
    public static ExcelWriterBuilder buildDownloadForHttpResponse(HttpServletResponse response, String fileName, boolean flag
            , List<WriteHandler>writeHandlerList) throws IOException {
        initResponse(response, fileName,flag);
        ExcelWriterBuilder write = EasyExcel.write(response.getOutputStream());
        writeHandlerList = Optional.ofNullable(writeHandlerList).orElse(new ArrayList<>());
        writeHandlerList.stream().forEach(writeHandler -> write.registerWriteHandler(writeHandler));
        write
        .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy())
        .needHead(true);
        if(flag){
            write.registerWriteHandler(SpringBeanUtil.getBean(SimpleCellWriteHandler.class))
                 .registerWriteHandler(SpringBeanUtil.getBean(SimpleSheetWriteHandler.class));
        }
        return write;
    }

    /**
     * 测试使用导出到本地文件
     * @param file
     * @param fileName
     * @param flag
     * @param writeHandlerList
     * @return
     * @throws IOException
     */
    public static ExcelWriterBuilder buildExcelWriterBuilder(File file, String fileName, boolean flag
            , List<WriteHandler>writeHandlerList) throws IOException {
        ExcelWriterBuilder write = EasyExcel.write(new FileOutputStream(file));
        writeHandlerList = Optional.ofNullable(writeHandlerList).orElse(new ArrayList<>());
        writeHandlerList.stream().forEach(writeHandler -> write.registerWriteHandler(writeHandler));
        write
                .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy())
                .needHead(true);
        if(flag){
            write.registerWriteHandler(SpringBeanUtil.getBean(SimpleCellWriteHandler.class))
                    .registerWriteHandler(SpringBeanUtil.getBean(SimpleSheetWriteHandler.class));
        }
        return write;
    }

    /**
     * 导入重新构建头，
     * @param analysisContext
     * @param headMap
     * @param clazz
     */
    public static void buildUpdateHeadAgain(AnalysisContext analysisContext, Map<Integer, ReadCellData<?>> headMap,Class clazz) {
        ExcelReadHeadProperty excelHeadPropertyData = analysisContext.readSheetHolder().excelReadHeadProperty();
        //获取导入的头
        Map<Integer, Head> nowHeadMapData = excelHeadPropertyData.getHeadMap();
        // 如果 nowHeadMapData 不为空，easyexcel能通过注解名字匹配上
        if (MapUtils.isNotEmpty(nowHeadMapData)) {
            return;
        }
        // 国际化处理将名字转换回中文重新匹配,originExcelHeadPropertyData由表头类产生
        ExcelReadHeadProperty originExcelHeadPropertyData = new ExcelReadHeadProperty(analysisContext.currentReadHolder(), clazz, null);
        //key下标，val列信息
        Map<Integer, Head> originHeadMapData = originExcelHeadPropertyData.getHeadMap();


        //excel实际列名
        Map<Integer, String> dataMap = ConverterUtils.convertToStringMap(headMap, analysisContext);
        Map<Integer, Head> tmpHeadMap = new HashMap<>(originHeadMapData.size() * 4 / 3 + 1);
        //临时类
        Map<Integer, ExcelContentProperty> tmpContentPropertyMap = new HashMap<>(originHeadMapData.size() * 4 / 3 + 1);
        //循环匹配
        for (Map.Entry<Integer, Head> entry : originHeadMapData.entrySet()) {
            Head headData = entry.getValue();
            String headName = String.format("%s.%s",headData.getField().getDeclaringClass().getSimpleName(),headData.getFieldName());
            headName = SpringBeanUtil.getBean(MessageSource.class).getMessage(headName,null, LocaleContextHolder.getLocale());
//            headData.setHeadNameList(Arrays.asList(headName));
            for (Map.Entry<Integer, String> stringEntry : dataMap.entrySet()) {
                if (stringEntry == null) {
                    continue;
                }
                String headString = stringEntry.getValue().trim();
                //下标
                Integer stringKey = stringEntry.getKey();
                if (StringUtils.isEmpty(headString)) {
                    continue;
                }

                if (StringUtils.equals(headName,headString)) {
                    headData.setColumnIndex(stringKey);
                    tmpHeadMap.put(stringKey, headData);
                    break;
                }
            }
        }
        excelHeadPropertyData.setHeadMap(tmpHeadMap);
    }

}
