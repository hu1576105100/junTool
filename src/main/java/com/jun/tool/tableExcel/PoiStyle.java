package com.jun.tool.tableExcel;

import lombok.Data;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.MessageFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.apache.poi.ss.usermodel.CellType.*;

@Data
public class PoiStyle{

    public XSSFWorkbook workbook;
    public Sheet sheet;

    private int needRow =0;//游标单元格所占用最大行数，用户确认下一次填充的行数

    private int nextRow=0;//游标行数
    private int nextCol=0;//游标列数

    private List<Cols> colWides= new ArrayList<>();//自动列宽

    public static DateTimeFormatter time = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    public static DateTimeFormatter date = DateTimeFormatter.ofPattern("yyyy年MM月dd日");

    public static String toString(LocalDateTime time, DateTimeFormatter data){
        return data.format(time);
    }

    public Map<Integer,Map<Integer,CellRangeAddress>> merges = new HashMap<>();

    public PoiStyle(XSSFWorkbook wb){
        this.workbook=wb;
//        this.sheet=workbook.getSheetAt(0);、、获取sheet页
        this.sheet=workbook.createSheet();//创建sheet页
    }

    public void initMerge(){
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress ca = sheet.getMergedRegion(i);
            if(ca.getLastColumn()-ca.getFirstColumn()>0||ca.getLastRow()-ca.getFirstRow()>0){
                Map<Integer,CellRangeAddress> newAddress = merges.get(ca.getFirstRow());
                if(isEmpty(newAddress)){
                    newAddress = new HashMap<>();
                    newAddress.put(ca.getFirstColumn(),ca);
                }else{
                    newAddress.put(ca.getFirstColumn(),ca);
                }
                merges.put(ca.getFirstRow(),newAddress);
            }
        }
    }

    public CellRangeAddress getMerge(Integer rowIndex,Integer colIndex){
        Map<Integer,CellRangeAddress> newAddress = merges.get(rowIndex);
        if(isEmpty(newAddress)){
            return null;
        }
        return newAddress.get(colIndex);
    }


    /**
     * 字体加粗
     */
    public void setBold(Cell cell){
        CellStyle cellStyle = cell.getCellStyle();
        //创建字体对象
        //注意，如果之前已经创建了字体则不能修改
        Font font = getFont(cellStyle);
        font.setBold(true);
        cellStyle.setFont(font);
        cell.setCellStyle(cellStyle);
    }
    /**
     * 字体颜色
     */
    public void setFontColor(Cell cell,IndexedColors color){
        CellStyle cellStyle = cell.getCellStyle();
        //获取字体对象
        Font font = getFont(cellStyle);
        font.setColor(color.getIndex());
        cellStyle.setFont(font);
        cell.setCellStyle(cellStyle);
    }

    /**
     * 字体换行
     */
    public void wrapText(Cell cell,boolean wrapText){
        CellStyle cellStyle = cell.getCellStyle();
        //是否自动换行
        cellStyle.setWrapText(wrapText);
        cell.setCellStyle(cellStyle);
    }

    /**
     * 背景颜色
     */
    public void setBgColor(Cell cell,HSSFColor.HSSFColorPredefined color){
        CellStyle style = cell.getCellStyle();
        if(style.getFillBackgroundColor()==64){
            style.setFillForegroundColor(color.getIndex());
        }
        cell.setCellStyle(style);
    }

    /**
     * 设置边框
     */
    public void setBorder(Cell cell){
        CellStyle cellStyle = cell.getCellStyle();
        cellStyle.setBorderBottom(BorderStyle.THIN); //下边框
        cellStyle.setBorderLeft(BorderStyle.THIN);//左边框
        cellStyle.setBorderTop(BorderStyle.THIN);//上边框
        cellStyle.setBorderRight(BorderStyle.THIN);//右边框、
        cell.setCellStyle(cellStyle);
    }

    /**
     * 设置边框（合并单元格）
     * @param firstRow
     * @param lastRow
     * @param firstCol
     * @param lastCol
     */
    public void setBorder(int firstRow, int lastRow, int firstCol, int lastCol){
        CellRangeAddress rangeAddress = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        RegionUtil.setBorderBottom(BorderStyle.THIN, rangeAddress, sheet);
        RegionUtil.setBorderLeft(BorderStyle.THIN, rangeAddress, sheet);
        RegionUtil.setBorderTop(BorderStyle.THIN, rangeAddress, sheet);
        RegionUtil.setBorderRight(BorderStyle.THIN, rangeAddress, sheet);
    }

    /**
     * 合并单元格
     */
    public void setMerge(int firstRow, int lastRow, int firstCol, int lastCol){
        CellRangeAddress rangeAddress = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        sheet.addMergedRegion(rangeAddress);
    }

    public void setStyle(Cell cell){

        //创建样式对象
        CellStyle style = cell.getCellStyle();
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cell.setCellStyle(style);
    }
    /**
     * 创建样式
     * @return
     */
    public CellStyle getStyle(){
        //创建样式对象
        CellStyle style = workbook.createCellStyle();
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    public void setAlignment(Cell cell,HorizontalAlignment alignment){
        //创建样式对象
        CellStyle style = cell.getCellStyle();
        //设置样式对齐方式：
        style.setAlignment(alignment);
        cell.setCellStyle(style);
    }
    public void setVertAlign(Cell cell,VerticalAlignment vertAlign){
        //创建样式对象
        CellStyle style = cell.getCellStyle();
        //设置样式对齐方式：
        style.setVerticalAlignment(vertAlign);
        //设定填充单色
        cell.setCellStyle(style);
    }


    private Font getFont(CellStyle cellStyle){
        if(cellStyle.getFontIndexAsInt()==0){
            return workbook.createFont();
        }else{
            return workbook.getFontAt(cellStyle.getFontIndexAsInt());
        }
    }

    /**
     * 根据行号获取当前行
     * @param rowNum
     * @return
     */
    public Row getRow(int rowNum){
        Row row = this.sheet.getRow(rowNum);
        if(isEmpty(row)){
            row = this.sheet.createRow(rowNum);
        }
        return row;
    }

    /**
     * 单行填充数据
     */
    public void fillData(Object data) throws IllegalAccessException {

        //初始化占用行数
        needRow=0;
        //填充当前行的值
        setCell(data,this.nextRow,nextCol);

        //设置下次数据开始行
        nextRow=1+needRow+nextRow;
    }


    public int setCell(Object data, int startRow, int startCol) throws IllegalAccessException {
        //获取当前行
        Row row = getRow(startRow);

        int size=1;//最少有一条数据,确定最大行数

        int title = 0;//判断是否有表头

        for (Field field:data.getClass().getDeclaredFields()) {
            field.setAccessible(true);
            if(isNotEmpty(field.get(data))&&field.getType().equals(List.class)){
                List objs = (List) field.get(data);
                if(objs.size()==0){
                    continue;
                }
                for (Field detailField:objs.get(0).getClass().getDeclaredFields()) {
                    TableExcel attr = detailField.getAnnotation(TableExcel.class);
                    if(isNotEmpty(attr.title())){
                        title = 1;
                    }
                }
                if(size<objs.size()) {
                    size = objs.size();
                }

            }
        }

        for (int i =0;i<data.getClass().getDeclaredFields().length;i++) {
            Field field = data.getClass().getDeclaredFields()[i];
            field.setAccessible(true);
            Object value = field.get(data);
            TableExcel attr = field.getAnnotation(TableExcel.class);

            //跳过该字段
            if(attr.ignore()){
                continue;
            }

            if(isNotEmpty(value)&&field.getType().equals(List.class)){
                List<Object> objs = (List<Object>) value;
                int startRow1 = startRow;
                int startCol1 = startCol;

                //列表先填写表头
                if(title==1){
                    setTitle(objs.get(0),startRow1,startCol1);
                    startRow1++;
                }

                for (Object objData : objs) {
                    startCol=setCell(objData,startRow1,startCol1);
                    startRow1++;
                }
                continue;
            }

            //获取字段注解值

            value = format(value,attr.value());

            //将该字段合并至前面一个单元格
            if(attr.mergeCol()!=-1){
                Cell cell = row.getCell(attr.mergeCol());
                cell.setCellValue(cell.getStringCellValue()+value);
                continue;
            }

            //单元格开始行号
            int firstRow;
            if(attr.row()!=-1){
                firstRow=attr.row();
            }else {
                firstRow=startRow;
            }

            //单元格结束行号
            int lastRow;
            if(attr.autoMergeRow()){
                lastRow=size+firstRow-1+title;
            }else{
                lastRow = attr.needRow()+firstRow-1;
            }

            //单元格开始列号
            int firstCol;
            if(attr.col()!=-1){
                firstCol = attr.col();
            }else{
                firstCol=startCol;
            }

            //单元格结束列号
            int lastCol = firstCol+attr.needCol()-1;




            //是否需要合并
            boolean merge=false;
            if(firstRow!=lastRow||firstCol!=lastCol){
                merge=true;
            }

            //合并单元格
            if(merge) {
                setMerge(firstRow, lastRow, firstCol, lastCol);
                //计算最大占用行数，目的：确认下次进入方法时的行数
                if(lastRow-firstRow>needRow){
                    needRow=lastRow-firstRow;
                }
            }

            //设置单元格的值
            Cell cell = row.createCell(firstCol);

            if(isEmpty(value)){
                cell.setCellValue("");
            }else if(field.getType().equals(String.class)){
                //设置单元格内容
                cell.setCellValue(String.valueOf(value));
                cell.setCellType(STRING);
            }else if(field.getType().equals(LocalDateTime.class)){
                cell.setCellValue(toString((LocalDateTime)value,date));
                cell.setCellType(STRING);
            }else if(field.getType().equals(Long.class)){
                cell.setCellValue((Long)value);
                cell.setCellType(NUMERIC);
            }else if(field.getType().equals(Double.class)){
                cell.setCellValue((Double)value);
                cell.setCellType(NUMERIC);
            }else if(field.getType().equals(Integer.class)){
                cell.setCellValue((Integer)value);
                cell.setCellType(NUMERIC);
            }else if(field.getType().equals(BigDecimal.class)){
                cell.setCellType(NUMERIC);
                double decimal = ((BigDecimal)value).doubleValue();
                cell.setCellValue(decimal);
            }else{
                cell.setCellType(STRING);
                cell.setCellValue(value.toString());
            }

            //设置边框
            if(merge) {
                setBorder(firstRow, lastRow, firstCol, lastCol);
                setStyle(cell);
            }else{
                cell.setCellStyle(getStyle());
                setBorder(cell);
            }

            setCellStyle(cell,attr);

            startCol=firstCol+attr.needCol();
        }


        return startCol;
    }




    public int setTitle(Object data, int startRow, int startCol) {
        //获取当前行
        Row row = getRow(startRow);

        for (Field field:data.getClass().getDeclaredFields()) {
            field.setAccessible(true);

            //获取字段注解值
            TableExcel attr = field.getAnnotation(TableExcel.class);

            //跳过该字段
            if(attr.mergeCol()!=-1||attr.ignore()){
                continue;
            }



            //单元格开始行号
            int firstRow;
            if(attr.row()!=-1){
                firstRow=attr.row();
            }else {
                firstRow=startRow;
            }

            //单元格结束行号
            int lastRow=firstRow;

            //单元格开始列号
            int firstCol;
            if(attr.col()!=-1){
                firstCol = attr.col();
            }else{
                firstCol=startCol;
            }

            //单元格结束列号
            int lastCol = firstCol+attr.needCol()-1;

            //是否需要合并
            boolean merge=false;
            if(firstRow!=lastRow||firstCol!=lastCol){
                merge=true;
            }

            //合并单元格
            if(merge) {
                setMerge(firstRow, lastRow, firstCol, lastCol);
                //计算最大占用行数，目的：确认下次进入方法时的行数
                if(lastRow-firstRow>needRow){
                    needRow=lastRow-firstRow;
                }
            }

            //设置单元格的值
            Cell cell = row.createCell(firstCol);

            cell.setCellValue(attr.title());

            //设置边框
            if(merge) {
                setBorder(firstRow, lastRow, firstCol, lastCol);
                setStyle(cell);
            }else{
                cell.setCellStyle(getStyle());
                setBorder(cell);
            }

            setBold(cell);
            setAlignment(cell,HorizontalAlignment.CENTER);
            setBgColor(cell,HSSFColor.HSSFColorPredefined.GREY_25_PERCENT);
            wrapText(cell,true);

            startCol=firstCol+attr.needCol();

            for (int i = 0; i < attr.virtualTitle();i++) {
                setVirtualTitle(row,startCol);
                startCol++;
            }
        }

        return startCol;
    }

    private void setVirtualTitle(Row row,int startCol){
        //设置单元格的值
        Cell cell = row.createCell(startCol);
        cell.setCellStyle(getStyle());
        setBorder(cell);
        setBold(cell);
        setBgColor(cell,HSSFColor.HSSFColorPredefined.GREY_25_PERCENT);

    }


    private void setCellStyle(Cell cell,TableExcel attr){
        //字体加粗
        if(attr.bold()) {
            setBold(cell);
        }
        //字体颜色
        setFontColor(cell,attr.fontColor());

        //设置背景色
        setBgColor(cell,attr.backGroundColor());

        //水平对齐方式
        setAlignment(cell,attr.setHorizAlign());

        //垂直对齐方式
        setVertAlign(cell,attr.setVertAlign());

        wrapText(cell,attr.wrapText());
    }


    private static  Object format(Object source,String format){
        if(isEmpty(format)){
            return source;
        }else if (isEmpty(source)){
            return MessageFormat.format(format,"");
        }else{
            return MessageFormat.format(format,source);
        }
    }

    /**
     * 根据行内容重新计算类宽
     */
    public void calcAndSetColwide() {
        //计算列宽
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row sourceRow=  sheet.getRow(i);
            for (int cellIndex = sourceRow.getFirstCellNum(); cellIndex < sourceRow.getPhysicalNumberOfCells(); cellIndex++) {
                Cols col = new Cols();
                if(colWides.size()>cellIndex){
                    col = colWides.get(cellIndex);
                }else{
                    colWides.add(col);
                }
                col.setCellIndex(cellIndex);


                if(isNotEmpty(getMerge(i,cellIndex))){
                    continue;
                }
                Cell cell = sourceRow.getCell(cellIndex);
                if(isEmpty(cell)){
                    continue;
                }
                int width = getWidth(getCellContentAsString(cell));
                if(width>col.getMaxLength()){
                    col.setMaxLength(width);
                }
            }
        }

//        //计算合并列宽
//        for(Integer rowIndex:merges.keySet()){
//            Map<Integer,CellRangeAddress> colArr = merges.get(rowIndex);
//            Row row = sheet.getRow(rowIndex);
//            for (Integer columnIndex : colArr.keySet()) {
//
//                Cell cell = row.getCell(columnIndex);
//                int width = getWidth(getCellContentAsString(cell));
//
//                CellRangeAddress addr = colArr.get(columnIndex);
//                int initWidth = 0;
//                if(colWides.size()<=addr.getFirstColumn()){
//                    Cols col = new Cols();
//                    col.setCellIndex(addr.getFirstColumn());
//                    col.setMaxLength(width);
//                    colWides.add(col);
//                    continue;
//                }
//
//                for(int i=addr.getFirstColumn();i<=addr.getLastColumn();i++){
//                    Cols col = colWides.get(i);
//                    initWidth+=col.getMaxLength();
//                }
//
//                if(width<=initWidth){
//                    continue;
//                }
//                int mergeSize = addr.getLastColumn() - addr.getFirstColumn() + 1;
//
//                int addWidth = (width-initWidth)/mergeSize;
//                for(int i=addr.getFirstColumn();i<=addr.getLastColumn();i++){
//                    Cols col = colWides.get(i);
//                    col.setMaxLength(col.getMaxLength()+addWidth);
//                }
//            }
//        }

        if(colWides.size() > 0){
            colWides.forEach(v->{
                sheet.setColumnWidth(v.getCellIndex(),v.getMaxLength());
            });
        }
    }


    //计算字符串列宽
    public static int getWidth(String str){
        char[] c = str.toCharArray();
        float length=0;
        for(int i = 0; i < c.length; i ++){
              String len = Integer.toBinaryString(c[i]);
              if(len.length() > 8){
                  length+=1.995;
              }else if(Character.isDigit(c[i])){
                  length++;
              }else{
                  length+=0.983;
              }

        }
        return (int)(length*277);
    }

    /**
     * 根据行内容重新计算行高
     */
    public void calcAndSetRowHigh() {
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row sourceRow = sheet.getRow(i);


            for (int cellIndex = sourceRow.getFirstCellNum(); cellIndex <= sourceRow.getPhysicalNumberOfCells(); cellIndex++) {
                //行高
                double maxHeight = sourceRow.getHeight();
                double stringNeedsRows=0;

                Cell sourceCell = sourceRow.getCell(cellIndex);
                //单元格的内容
                String cellContent = getCellContentAsString(sourceCell);

                if (isEmpty(cellContent)) {continue; }

                //单元格的宽高及单元格信息
                Map<String, Object> cellInfoMap = getCellInfo(sourceCell);
                Integer cellWidth = (Integer) cellInfoMap.get("width");
                Integer cellHeight = (Integer) cellInfoMap.get("height");
                if (cellHeight > maxHeight) {
                    maxHeight = cellHeight;
                }
//            System.out.println("单元格的宽度 : " + cellWidth + "    单元格的高度 : " + maxHeight + ",    单元格的内容 : " + cellContent);
                CellStyle cellStyle = sourceCell.getCellStyle();
                Font font = getFont(cellStyle);
                //字体的高度
                short fontHeight = font.getFontHeight();
                //cell内容字符串总宽度
//                double cellContentWidth = cellContent.getBytes().length * 2 * 256;
                double cellContentWidth = getWidth(cellContent);

                //字符串需要的行数 不做四舍五入之类的操作
                stringNeedsRows = Math.ceil(cellContentWidth / cellWidth);
                //小于一行补足一行


                //需要的高度 			(Math.floor(stringNeedsRows) - 1) * 40 为两行之间空白高度
                double stringNeedsHeight = (double) fontHeight * stringNeedsRows;
                //需要重设行高
                if (stringNeedsHeight > maxHeight) {
                    maxHeight = stringNeedsHeight;
                    //超过原行高三倍 则为5倍 实际应用中可做参数配置
                    if (maxHeight / cellHeight > 5) {
                        maxHeight = 5 * cellHeight;
                    }
                    //最后取天花板防止高度不够
                    maxHeight = Math.ceil(maxHeight);
                    //重新设置行高 同时处理多行合并单元格的情况
                    Boolean isPartOfRowsRegion = (Boolean) cellInfoMap.get("isPartOfRowsRegion");
                    if (isPartOfRowsRegion) {
                        Integer firstRow = (Integer) cellInfoMap.get("firstRow");
                        Integer lastRow = (Integer) cellInfoMap.get("lastRow");
                        //平均每行需要增加的行高
                        double addHeight = (maxHeight - cellHeight) / (lastRow - firstRow + 1);
                        for (int j = firstRow; j <= lastRow; j++) {
                            double rowsRegionHeight = sourceRow.getSheet().getRow(j).getHeight() + addHeight;
                            sourceRow.getSheet().getRow(j).setHeight((short) rowsRegionHeight);
                        }
                    } else {
                        sourceRow.setHeight((short) maxHeight);
                    }
                }
            }
        }
    }


    /**
     * 解析一个单元格得到数据
     * @param cell
     * @return
     */
    private static String getCellContentAsString(Cell cell) {
        if(null == cell){
            return "";
        }
        String result = "";
        switch (cell.getCellType()) {
            case NUMERIC:
                String s = String.valueOf(cell.getNumericCellValue());
                if (s != null) {
                    if (s.endsWith(".0")) {
                        s = s.substring(0, s.length() - 2);
                    }
                }
                result = s;
                break;
            case STRING:
                result = String.valueOf(cell.getStringCellValue()).trim();
                break;
            case BLANK:
                break;
            case BOOLEAN:
                result = String.valueOf(cell.getBooleanCellValue());
                break;
            case ERROR:
                break;
            default:
                break;
        }
        return result;
    }


    /**
     * 获取单元格及合并单元格的宽度
     * @param cell
     * @return
     */
    private Map<String, Object> getCellInfo(Cell cell) {
        Sheet sheet = cell.getSheet();
        int rowIndex = cell.getRowIndex();
        int columnIndex = cell.getColumnIndex();

        boolean isPartOfRegion = false;
        int firstColumn = 0;
        int lastColumn = 0;
        int firstRow = 0;
        int lastRow = 0;


        CellRangeAddress ca  = getMerge(rowIndex,columnIndex);
        if(isNotEmpty(ca)){
            firstColumn = ca.getFirstColumn();
            lastColumn = ca.getLastColumn();
            firstRow = ca.getFirstRow();
            lastRow = ca.getLastRow();
            isPartOfRegion = true;
        }

        Map<String, Object> map = new HashMap<String, Object>();
        Integer width = 0;
        Integer height = 0;
        boolean isPartOfRowsRegion = false;
        if(isPartOfRegion){
            for (int i = firstColumn; i <= lastColumn; i++) {
                width += sheet.getColumnWidth(i);
            }
            for (int i = firstRow; i <= lastRow; i++) {
                height += sheet.getRow(i).getHeight();
            }
            if(lastRow > firstRow){
                isPartOfRowsRegion = true;
            }
        }else{
            width = sheet.getColumnWidth(columnIndex);
            height += cell.getRow().getHeight();
        }

        map.put("isPartOfRowsRegion", isPartOfRowsRegion);
        map.put("firstRow", firstRow);
        map.put("lastRow", lastRow);
        map.put("width", width);
        map.put("height", height);
        return map;
    }


    /**
     * 判空
     */
    public static boolean isEmpty(Object obj){
        if(obj==null){
            return true;
        }
        if (obj instanceof Map) {
            return ((Map<?, ?>) obj).isEmpty();
        } else if (obj instanceof List) {
            return ((List< ? >) obj).isEmpty();
        } else if (obj instanceof String) {
            return ((String) obj).length() == 0;
        }
        return false;
    }
    public static boolean isNotEmpty(Object obj){
        return !isEmpty(obj);
    }
}
