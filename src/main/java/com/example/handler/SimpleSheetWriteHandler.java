package com.example.handler;


import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.context.i18n.LocaleContextHolder;
import org.springframework.stereotype.Component;

@Component
public class SimpleSheetWriteHandler implements SheetWriteHandler {
    @Autowired
    private MessageSource messageSource;
    @Override
    public void beforeSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        String sheetName = writeSheetHolder.getSheetName();
        String message = messageSource.getMessage(sheetName, null, LocaleContextHolder.getLocale());
        writeSheetHolder.setSheetName(message);
    }
}
