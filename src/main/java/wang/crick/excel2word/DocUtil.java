package wang.crick.excel2word;

import freemarker.template.Configuration;
import freemarker.template.DefaultObjectWrapper;
import freemarker.template.Template;
import freemarker.template.TemplateExceptionHandler;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;

public class DocUtil {
    private Configuration configure = null;

    public DocUtil() {
        configure = new Configuration();
        configure.setDefaultEncoding("utf-8");
    }

    /**
     * 根据Doc模板生成word文件
     *
     * @param dataMap  Map 需要填入模板的数据
     * @param savePath 保存路径
     */
    public void createDoc(Map<String, String> dataMap, String downloadType, String savePath) {
        downloadType = "template.xml";
        //getDataMap(dataMap);
        savePath = "./" + dataMap.get("name") + ".doc";
        try {
            //加载需要装填的模板
            Template template = null;
            //加载模板文件
            configure.setDirectoryForTemplateLoading(new File("./"));
            //设置对象包装器
            configure.setObjectWrapper(new DefaultObjectWrapper());
            //设置异常处理器
            configure.setTemplateExceptionHandler(TemplateExceptionHandler.IGNORE_HANDLER);
            //定义Template对象,注意模板类型名字与downloadType要一致
//			Thread.currentThread().getContextClassLoader().getResource("template.xml")
            template = configure.getTemplate(downloadType);
            //输出文档
            File outFile = new File(savePath);
            Writer out = null;
            out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFile), "utf-8"));
            template.process(dataMap, out);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) throws Exception {
        DocUtil dou = new DocUtil();
        String path = "./template.xls";
        InputStream in = new FileInputStream(path);
        Workbook book = new HSSFWorkbook(in);
        Sheet sheet = book.getSheetAt(0);
        for (int j = sheet.getFirstRowNum() + 1; j < sheet.getLastRowNum(); j++) {
            Row row = sheet.getRow(j);
            if (null == row) {
                break;
            }
            if (null == row.getCell(0)) {
                break;
            }
            Map<String, String> dataMap = new HashMap<String, String>();
            dataMap.put("name", row.getCell(0).toString());
            dataMap.put("gender", row.getCell(1).toString());
            dataMap.put("tel", row.getCell(2).toString());
            dataMap.put("nation", row.getCell(4).toString());
            String idcard = row.getCell(3).toString();
            String y = null;
            String m = null;
            String birthday = null;
            if (Objects.nonNull(idcard)) {
                if (idcard.length() == 15) {
                    y = "19" + idcard.substring(6, 8);
                    m = idcard.substring(8, 10);
                } else {
                    y = idcard.substring(6, 10);
                    m = idcard.substring(10, 12);
                }
                birthday = y + "." + Integer.parseInt(m);
            }
            dataMap.put("birthday", birthday);
            dataMap.put("address", row.getCell(5).toString());
            dataMap.put("idcard", row.getCell(3).toString());
            dou.createDoc(dataMap, "", "");
        }
    }
}
