package com.example.handler;


import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.context.i18n.LocaleContextHolder;
import org.springframework.stereotype.Component;
import sun.misc.resources.Messages;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

@Component
public class SimpleCellWriteHandler implements CellWriteHandler {
    @Autowired
    private MessageSource messageSource;
    @Override
    public void beforeCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Head head, Integer columnIndex, Integer relativeRowIndex, Boolean isHead) {
        //只处理头部
        if(isHead){
            String simpleName = head.getField().getDeclaringClass().getSimpleName();
            String fieldName = head.getFieldName();
            String message = messageSource.getMessage(String.format("%s.%s", simpleName, fieldName), null, LocaleContextHolder.getLocale());
            head.setHeadNameList(Arrays.asList(message));
        }

    }
}
