package com.jun.tool;

import com.jun.tool.exception.BusinessExceptionAssert;

public enum SysTemAssert implements BusinessExceptionAssert {

    SourceNotNull(500, "Source must not be null"),
    TargetNotNull(500, "Target must not be null"),
    ;
    /**
     * 返回码
     */
    private Integer code;
    /**
     * 返回消息
     */
    private String message;

    SysTemAssert(Integer code, String message) {
        this.code = code;
        this.message = message;
    }

    @Override
    public Integer getCode() {
        return code;
    }

    @Override
    public String getMessage() {
        return message;
    }
}
