package com.example.handler;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.data.ReadCellData;
import com.alibaba.excel.read.metadata.holder.ReadRowHolder;
import com.example.util.ExcelUtil;
import com.example.vo.FillData;

import java.util.Map;

public class SimpReadleListener extends AnalysisEventListener<FillData> {


    @Override
    public void invokeHead(Map<Integer, ReadCellData<?>> headMap, AnalysisContext context) {
        ReadRowHolder readRowHolder = context.readRowHolder();
        int rowIndex = readRowHolder.getRowIndex();
        int currentHeadRowNumber = context.readSheetHolder().getHeadRowNumber();
        if (currentHeadRowNumber == rowIndex + 1) {
            //表头触发 表头国际化转换
            ExcelUtil.buildUpdateHeadAgain(context,headMap,FillData.class);
        }
    }

    @Override
    public void invoke(FillData fillData, AnalysisContext analysisContext) {
        System.out.println(fillData);
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {

    }
}
