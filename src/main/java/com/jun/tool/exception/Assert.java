package com.jun.tool.exception;

public interface Assert {
    /**
     * 创建异常
     */
    BaseException newException();

    BaseException newException(Throwable e);

    /**
     * 如果条件为真，则抛出异常
     */
    default void exception(Boolean decide) {
        if (decide) {
            throw newException();
        }
    }

    default void exception(Throwable e) {
        throw newException(e);
    }

    default void exception() {
        throw newException();
    }
}
