package com.jun.tool.tableWord;

import lombok.AllArgsConstructor;
import lombok.Getter;

@Getter
@AllArgsConstructor
public enum ContentType {
    Word_docx("application/vnd.openxmlformats-officedocument.wordprocessingml.document",".docx"),
    Word_doc("application/msword",".doc"),
    Excel_xlsx("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",".xlsx"),
    Excel_xls("application/vnd.ms-excel",".xls"),
    PPT("application/vnd.ms-powerpoint",".ppt"),
    ;
    private final String type;
    private final String suffix;
}
