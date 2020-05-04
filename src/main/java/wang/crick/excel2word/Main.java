package wang.crick.excel2word;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.List;
import java.util.Map;

import freemarker.template.Template;
import wang.crick.excel2word.config.constants.WorkConstant;
import wang.crick.excel2word.config.enumerations.FileType;
import wang.crick.excel2word.template.TemplateManager;
import wang.crick.excel2word.utils.XlsUtil;
import wang.crick.excel2word.utils.DocUtil;

public class Main {

    public static void main(String[] args) {
        Template template = null;
        FileInputStream xlsFin = null;
        Writer out = null;
        List<Map<String, Object>> xlsList = null;
        try {
            template = TemplateManager.loading(FileType.XML_TYPE);
            xlsFin = new FileInputStream(new File(WorkConstant.WORK_PATH+WorkConstant.TEMPLATE_NAME+FileType.EXCEL_03_TYPE.getLabel()));
            xlsList = XlsUtil.extractXls(xlsFin);
            for (Map<String, Object> xls : xlsList) {
                //默认第一列为doc文件名
                File doc = DocUtil.create(FileType.WORD_03_TYPE, xls.get("name").toString());
                out= new BufferedWriter(new OutputStreamWriter(new FileOutputStream(doc),"utf-8"));
                template.process(xls, out);
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            TemplateManager.destory();
            if (null != xlsFin) {
                try {
                    xlsFin.close();
                } catch (IOException e) {
                    e.printStackTrace();
                } finally {
                    xlsFin = null;
                }
            }
        }
    }
}
