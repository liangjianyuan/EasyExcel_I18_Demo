package com.example.vo;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.DateTimeFormat;
import lombok.*;

import java.util.Date;

@Getter
@Setter
@EqualsAndHashCode
@AllArgsConstructor
@NoArgsConstructor
public class FillData {
    @ExcelProperty({"姓名"})
    private String name;
    @ExcelProperty({"手机号"})
    private String phone;
    @ExcelProperty({"生日"})
    @DateTimeFormat("yyyy-MM")
    private Date date;
}