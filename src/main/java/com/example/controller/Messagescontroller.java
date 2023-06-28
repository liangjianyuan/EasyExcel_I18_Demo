package com.example.controller;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.style.column.LongestMatchColumnWidthStyleStrategy;
import com.example.util.ExcelUtil;
import com.example.vo.FillData;
import com.sun.deploy.net.HttpResponse;
import lombok.SneakyThrows;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.context.i18n.LocaleContextHolder;
import org.springframework.http.HttpRequest;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;

@RestController
public class Messagescontroller {

    @Autowired
    MessageSource messageSource;

    @GetMapping("/test")
    public String test(){
        return messageSource.getMessage("lang.FillData.name", null, LocaleContextHolder.getLocale());
    }


    @GetMapping("/download/temple")
    public void download(HttpServletResponse response) throws IOException {
        ExcelWriterBuilder build = ExcelUtil.buildDownloadForHttpResponse(response, "FillData.fileName", true, null);
        ExcelWriter writer = build.build();
        writer.write(Arrays.asList(new FillData("zhangsan","1310000000",new Date()))
                , EasyExcel.writerSheet(1,"FillData.fileName").head(FillData.class).build());
        writer.finish();
    }

    @GetMapping("/download/easy")
    public void downloadEasy(HttpServletResponse response) throws IOException {
        // 这里注意 有同学反应使用swagger 会导致各种问题，请直接用浏览器或者用postman
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setCharacterEncoding("utf-8");
        // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系
        String fileName = URLEncoder.encode("测试", "UTF-8").replaceAll("\\+", "%20");
        response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
        EasyExcel.write(response.getOutputStream(), FillData.class)
                .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy())
                .sheet("模板").doWrite(new ArrayList<>());
    }
}
