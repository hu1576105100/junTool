package com.jun.tool.tableWord;

import lombok.Data;

import java.util.List;

/**
 * 富文本片段
 */
@Data
public class WordRTF {

    /**
     * 标题
     */
    private String title;

    /**
     * 富文本
     */
    private String value;

    /**
     * 子文本
     */
    private List<WordRTF> wordRTFS;
}
