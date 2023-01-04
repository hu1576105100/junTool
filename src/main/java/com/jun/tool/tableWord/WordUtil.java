package com.jun.tool.tableWord;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.OutlineLevel;
import com.aspose.words.SaveFormat;
import com.jun.tool.BeanUtil;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

public class WordUtil {

    public static void createWord(HttpServletResponse response, List<WordText> data, String titleName,Integer size) throws Exception {
        //创建文本对象
        XWPFDocument docxDocument = new XWPFDocument();

        //创建标题
        setTitle(titleName,docxDocument);

        //写入文本
        setValue(data,docxDocument,size);

        //获取输出流
        OutputStream  out = getOutputStream(response,titleName,ContentType.Word_docx);

        //写入文件
        docxDocument.write(out);

        out.close();

        docxDocument.close();
    }

    private static void setTitle(String titleName,XWPFDocument docxDocument){
        XWPFParagraph firstParagraphX = docxDocument.createParagraph();
        firstParagraphX.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun runTitle = firstParagraphX.createRun();
        runTitle.setText(titleName);
        runTitle.setBold(true);
        runTitle.setFontSize(24);
        runTitle.setFontFamily("宋体");
        runTitle.addCarriageReturn();//回车键
        runTitle.setKerning(30);
    }


    /**
     * 写入文本内容
     * @param data
     * @param docxDocument
     */
    private static void setValue(List<WordText> data,XWPFDocument docxDocument,Integer size) throws Exception {
        for (int j = 0; j <data.size(); j++){
            WordText wordDTO=data.get(j);
            XWPFParagraph paragraphX = docxDocument.createParagraph();
//            paragraphX.setFirstLineIndent(400);//首行缩进
            //创建段落中的run
            XWPFRun run = paragraphX.createRun();
            run.setFontSize(10);
            run.setText(wordDTO.getValue());
            run.addCarriageReturn();//回车键
            if(BeanUtil.isEmpty(size)){
                continue;
            }
            double page = getPage(docxDocument);
            if(page>size){
                break;
            }
        }
    }

    private static double getPage(XWPFDocument docxDocument) throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        docxDocument.write(out);
        ByteArrayInputStream input = new ByteArrayInputStream(out.toByteArray());
        Document document = new Document(input);
        int pages =document.getPageCount();
        input.close();
        out.close();
        return pages*2.1d;
    }

    /**
     * 导出富文本内容到word
     * @param response
     * @param data 输出内容
     * @param fileName 导出文件名称
     * @throws Exception
     */
    public static void exportHtmlToWord(HttpServletResponse response, List<WordRTF> data, String fileName) throws Exception {

        //转化富文本内容
        byte[] contents = getContents(data);

        //转输入流
        ByteArrayInputStream input = new ByteArrayInputStream(contents);

        //创建文档实体
        Document document = new Document(input);//首次调用初始化耗时1.553s

        //提供修改文档内容方法
        DocumentBuilder builder = new DocumentBuilder(document);

        //移动光标至文档开始
        builder.moveToDocumentStart();

        //生成标题
        title(builder,fileName);

        //生成目录(所有文本内容都写入最后生成目录)
        directory(builder);

        //跟新目录
        document.updateFields();//首次调用初始化耗时1.078s

        //获取输出流
        ServletOutputStream oStream = getOutputStream(response,fileName,ContentType.Word_docx);

        //格式化保存
        document.save(oStream, SaveFormat.DOCX);//首次调用初始化耗时0.205s

        oStream.close();

    }


    private static byte[] getContents(List<WordRTF> data) throws IOException {
        // 拼接html格式内容
        StringBuilder sbf = new StringBuilder();
        // 这里拼接一下html标签,便于word文档能够识别
        sbf.append("<html " +
                "xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:w=\"urn:schemas-microsoft-com:office:word\" xmlns:m=\"http://schemas.microsoft.com/office/2004/12/omml\" xmlns=\"http://www.w3.org/TR/REC-html40\"" + //将版式从web版式改成页面试图
                ">"+
                "<head>" +
                "<!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View><w:TrackMoves>false</w:TrackMoves><w:TrackFormatting/><w:ValidateAgainstSchemas/><w:SaveIfXMLInvalid>false</w:SaveIfXMLInvalid><w:IgnoreMixedContent>false</w:IgnoreMixedContent><w:AlwaysShowPlaceholderText>false</w:AlwaysShowPlaceholderText><w:DoNotPromoteQF/><w:LidThemeOther>EN-US</w:LidThemeOther><w:LidThemeAsian>ZH-CN</w:LidThemeAsian><w:LidThemeComplexScript>X-NONE</w:LidThemeComplexScript><w:Compatibility><w:BreakWrappedTables/><w:SnapToGridInCell/><w:WrapTextWithPunct/><w:UseAsianBreakRules/><w:DontGrowAutofit/><w:SplitPgBreakAndParaMark/><w:DontVertAlignCellWithSp/><w:DontBreakConstrainedForcedTables/><w:DontVertAlignInTxbx/><w:Word11KerningPairs/><w:CachedColBalance/><w:UseFELayout/></w:Compatibility><w:BrowserLevel>MicrosoftInternetExplorer4</w:BrowserLevel><m:mathPr><m:mathFont m:val=\"Cambria Math\"/><m:brkBin m:val=\"before\"/><m:brkBinSub m:val=\"--\"/><m:smallFrac m:val=\"off\"/><m:dispDef/><m:lMargin m:val=\"0\"/> <m:rMargin m:val=\"0\"/><m:defJc m:val=\"centerGroup\"/><m:wrapIndent m:val=\"1440\"/><m:intLim m:val=\"subSup\"/><m:naryLim m:val=\"undOvr\"/></m:mathPr></w:WordDocument></xml><![endif]-->" +
                "</head>" +
                "<body>");

        // 富文本内容
        sbf.append(getValueByHtml(data,null,1));

        sbf.append("</body></html>");
        // 必须要设置编码,避免中文就会乱码
        return sbf.toString().getBytes("utf-8");
    }

    private static StringBuffer getValueByHtml(List<WordRTF> data, Integer i, int fontsize) throws IOException {
        String orderNo="";
        if(i!=null){
            orderNo=(i+1)+".";
        }
        StringBuffer sbf = new StringBuffer();
        for (int j = 0; j <data.size(); j++){
            WordRTF wordRTF =data.get(j);
            sbf.append("<h"+fontsize+"><strong style=\"color: rgb(22, 22, 22);\">");
            sbf.append(orderNo+(j+1)+"."+ wordRTF.getTitle());
            sbf.append("</strong></h"+fontsize+">");
            if(BeanUtil.isNotEmpty(wordRTF.getValue())){
                sbf.append(wordRTF.getValue());
            }
            if(BeanUtil.isNotEmpty(wordRTF.getWordRTFS())){
                sbf.append(getValueByHtml(wordRTF.getWordRTFS(),j,fontsize+1));
            }
        }
        return sbf;
    }

    public static ServletOutputStream getOutputStream( HttpServletResponse response,String fileName,ContentType contentType) throws IOException {
        response.setContentType(contentType.getType());
        response.addHeader("Content-Disposition", "attachment;filename=" +
                new String(fileName.getBytes("utf-8"),"iso8859-1") + contentType.getSuffix());
        return response.getOutputStream();
    }

    private static void title (DocumentBuilder builder,String fileName){
        try {
            builder.getFont().setBold(true);//字体变粗
            builder.getFont().setSize(24);//字体大小
            builder.getParagraphFormat().setAlignment(1);//居中对其设置 Left:0 \Center:1 Right：2
            builder.getParagraphFormat().setOutlineLevel(OutlineLevel.LEVEL_1);
            builder.write(fileName);
            builder.insertParagraph();//插入分段符
            builder.getFont().setBold(false);
            builder.getFont().setSize(16);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void directory(DocumentBuilder builder){
        try {
            builder.getFont().setBold(false);//字体不变粗
            builder.getFont().setSize(16);//字体大小
            builder.getParagraphFormat().setAlignment(1);//居中对其设置 Left:0 \Center:1 Right：2
            builder.getParagraphFormat().setOutlineLevel(OutlineLevel.LEVEL_1);
            builder.write("目录");
            builder.insertParagraph();//插入分段符
            builder.insertTableOfContents("\\o \"1-4\" \\h \\z \\u");//添加4级目录
            builder.insertParagraph();//插入分段符
            builder.getParagraphFormat().setAlignment(ParagraphAlignment.RIGHT.getValue());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
