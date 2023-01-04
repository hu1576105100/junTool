package com.jun.tool.tableExcel;

import com.jun.tool.exception.BusinessExceptionAssert;
import lombok.AllArgsConstructor;
import lombok.Getter;

@Getter
@AllArgsConstructor
public enum ExcelAssert implements BusinessExceptionAssert {

    exportError(1001, "列表导出失败!"),
    ;
    /**
     * 返回码
     */
    private Integer code;
    /**
     * 返回消息
     */
    private String message;
}
