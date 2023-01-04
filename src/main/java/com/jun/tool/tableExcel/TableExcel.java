package com.jun.tool.tableExcel;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 自定义导出Excel数据注解
 * 
 * @author hujun
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface TableExcel
{

    /**
     * 指定列号
     */
    public int col() default -1;

    /**
     * 指定行号
     */
    public int row() default -1;

    /**
     * 需要列数
     */
    public int needCol() default 1;

    /**
     * 需要行数
     */
    public int needRow() default 1;


    /**
     * 自动合并当前行
     */
    public boolean autoMergeRow() default false;


    /**
     * 加粗
     */
    public boolean bold() default false;

    /**
     * 列向表头
     * 设置了该值后，列表数据向下一一行
     */
    public String title() default "";

    /**
     * 文字换行
     * 设置了该值后，超出单元格宽度后会自动换行
     */
    public boolean wrapText() default true;


    /**
     * 填充值 例：XXX是中国的   则value="{0}是中国的"
     * 注：设置该值后该值中没有占位符，则输出该值，不输出字段值
     */
    public String value() default "";

    /**
     * 字体颜色  获取方式 例：IndexedColors.RED.getIndex()
     */
    public IndexedColors fontColor() default IndexedColors.BLACK1;

    /**
     * 背景颜色  获取方式 例：IndexedColors.RED.getIndex()
     */
    public HSSFColor.HSSFColorPredefined backGroundColor() default HSSFColor.HSSFColorPredefined.WHITE;

    /**
     * 字体水平对齐方式  获取方式 例：HorizontalAlignment.CENTER
     */
    public HorizontalAlignment setHorizAlign() default HorizontalAlignment.CENTER;

    /**
     * 字体水平对齐方式  获取方式 例：VerticalAlignment.CENTER
     */
    public VerticalAlignment setVertAlign() default VerticalAlignment.CENTER;

    /**
     * 忽略该字段
     */
    public boolean ignore() default false;

    /**
     * 将字段合并到另一个单元格
     */
    public int mergeCol() default -1;


    /**
     * 单元格后面需要几个空格表头 （用于表头排版）
     */
    public int virtualTitle() default 0;

}