package com.jun.tool.tableExcel;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.util.List;


@Slf4j
public  class ExcelUtil {

    /**
     * 导出excel
     * @param dataList
     * @param response
     */
    public static void exportTable(List<Object> dataList, HttpServletResponse response,String name) {

        // 从XX行开始为数据内容  excel 第一行为0

        // 　2003版本的Excel （xls） ---- HSSFWorkbook
        //    2007版本以及更高版本 (xlsx)---- XSSFWorkbook
        XSSFWorkbook workbook = null;
        OutputStream out = null;
        try {

            //这个excel.path 是配置 上面模板的路径
            response.setContentType("application/vnd.ms-excel");
            response.setCharacterEncoding("utf-8");
            //导出的文件名
            String fileName = URLEncoder.encode( name+".xlsx", "utf-8");
            response.setHeader("Content-disposition", "attachment; filename=" + new String(fileName.getBytes("UTF-8"), "ISO-8859-1"));
            out = response.getOutputStream();


            workbook = new XSSFWorkbook();

            //自动义样式
            PoiStyle poiStyle = new PoiStyle(workbook);

            // 填充数据
            for(Object data:dataList){
                poiStyle.fillData(data);
            }
            poiStyle.initMerge();//初始化 合并单元格

            //计算列宽
            poiStyle.calcAndSetColwide();
            //计算行高
            poiStyle.calcAndSetRowHigh();
            // 输出流
            // excel工作空间写入流
            workbook.write(out);
            out.flush();
        } catch (Exception e) {
            ExcelAssert.exportError.exception(e);
        } finally {
            try {
                workbook.close();
                out.close();
            } catch (IOException e) {
                ExcelAssert.exportError.exception(e);
            }
        }
    }
}
