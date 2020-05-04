package wang.crick.excel2word.template;

import freemarker.template.Configuration;
import freemarker.template.DefaultObjectWrapper;
import freemarker.template.Template;
import freemarker.template.TemplateExceptionHandler;
import wang.crick.excel2word.config.constants.WorkConstant;
import wang.crick.excel2word.config.enumerations.FileType;

import java.io.*;
import java.util.Objects;

public class TemplateManager {

    private static Configuration config = null;

    private static String templatePath;

    /**
     * 生成模板temp文件
     *
     * @param TemplateType type 模板类型
     * @return Template 模板
     * @throws IOException
     */
    public static Template loading(FileType templateType) throws IOException {
        config = new Configuration();
        config.setDefaultEncoding("utf-8");
        //TODO work目录不存在 可自动创建
        Template template  = null;
        InputStream in = null;
        BufferedReader reader = null;
//		Writer out = null;
        FileOutputStream fos = null;
        config.setObjectWrapper(new DefaultObjectWrapper());
        config.setTemplateExceptionHandler(TemplateExceptionHandler.IGNORE_HANDLER);
        try {
            config.setDirectoryForTemplateLoading(new File(WorkConstant.WORK_PATH) );
            in = new FileInputStream(new File(WorkConstant.WORK_PATH
                    + WorkConstant.TEMPLATE_NAME + templateType.getLabel()));
            reader = new BufferedReader(new InputStreamReader(in));
            String line = null;
            StringBuilder sb = new StringBuilder();
            while ((line = reader.readLine()) != null) {
                sb.append(line);
            }
            TemplateManager.templatePath = WorkConstant.WORK_PATH
                    + WorkConstant.TEMP_NAME + templateType.getLabel();
//			out = new BufferedWriter(new FileWriter(templatePath));
//			out.write(sb.toString().replaceAll("crmark0?(\\d{1,2})", "\\${m$1}"));

            fos = new FileOutputStream(new File(templatePath));
            fos.write(sb.toString().replaceAll("crmark0?(\\d{1,2})", "\\${$1}").getBytes());
            template = config.getTemplate(WorkConstant.TEMP_NAME + templateType.getLabel());
        } catch (IOException e) {
            e.printStackTrace();
            destory();
            throw e;
        } finally {
            try {
                if (fos != null) {
                    fos.close();
                }
//				if (out != null) {
//					out.close();
//				}
                if (in != null) {
                    in.close();
                }
                if (reader != null) {
                    reader.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            } finally {
//				out = null;
                fos = null;
                in = null;
                reader = null;
            }
        }
        return template;
    }

    /**
     * 销毁模板，清除temp文件
     */
    public static void destory() {
        if(Objects.nonNull(TemplateManager.templatePath)) {
            File file = new File(TemplateManager.templatePath);
            if (file.exists()) {
                file.delete();
            }
            TemplateManager.templatePath = null;
            config.clearTemplateCache();
        }
    }
}