package com.example;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.example.handler.SimpReadleListener;
import com.example.util.ExcelUtil;
import com.example.vo.FillData;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.context.i18n.LocaleContextHolder;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Locale;


@SpringBootTest
@RunWith(SpringRunner.class)
public class MessagesTest {

    @Test
    public void uploadTest() throws IOException {
        //测试导入日文表头
        LocaleContextHolder.setLocale(Locale.JAPAN);
        String fileName = "D://ユーザデータ .xlsx";
        //在读取时重新构建表头
        EasyExcel.read(fileName, FillData.class, new SimpReadleListener()).sheet().doRead();
    }

    @Test
    public void downloadTest() throws  Exception{
        //下载日本模板
        String fileName = "D://ユーザデータ .xlsx";
        LocaleContextHolder.setLocale(Locale.JAPAN);
        ExcelWriterBuilder excelWriterBuilder = ExcelUtil.buildExcelWriterBuilder(new File(fileName), "", true, null);
        //对sheet设置头
        excelWriterBuilder.build().write(new ArrayList<>(),EasyExcel.writerSheet("人员数据").head(FillData.class).build()).close();

    }


}
