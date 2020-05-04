package wang.crick.excel2word.utils;

import wang.crick.excel2word.config.constants.WorkConstant;
import wang.crick.excel2word.config.enumerations.FileType;

import java.io.File;


public class DocUtil {

    public static File create(FileType type, String docName) {
        //解决重名问题
        int i = 1;
        String outPath = WorkConstant.WORK_PATH + docName+type.getLabel();
        File doc = new File(outPath);
        while (doc.exists()) {
            outPath = WorkConstant.WORK_PATH + docName+"("+i+").doc";
            doc = new File(outPath);
            i++;
        }
        return doc;
    }
}
