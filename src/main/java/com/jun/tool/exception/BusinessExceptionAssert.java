package com.jun.tool.exception;

public interface BusinessExceptionAssert extends IResponseEnum, Assert {

    @Override
    default BaseException newException() {
        return new BaseException(this.getMessage(),this.getCode());
    }

    @Override
    default BaseException newException(Throwable e){
        return new BaseException(this.getMessage(),this.getCode(),e);
    }
}
