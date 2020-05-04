package wang.crick.excel2word.utils;


import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class XlsUtil {

    public static Map<String, Object> extractRow(Row row, List<String> keys){
        Map<String, Object> params = new HashMap<String, Object>();
        for (int i=0; i<row.getLastCellNum(); i++) {
            params.put("name", row.getCell(0).toString());
            params.put("gender", row.getCell(1).toString());
            params.put("tel", row.getCell(2).toString());
            params.put("nation", row.getCell(4).toString());
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
            params.put("birthday", birthday);
            params.put("address", row.getCell(5).toString());
            params.put("idcard", row.getCell(3).toString());
        }
        return params;
    }

    /**
     * 第一行作为key，即mark标签名
     */
    private static List<String> getKeyList(Row row) {
        List<String> keys = new ArrayList<String>();
        for (int i=0; i<=row.getLastCellNum();i++) {
            if (row.getCell(i) != null) {
                row.getCell(i).setCellType(Cell.CELL_TYPE_STRING);
            }
            keys.add(row.getCell(i)==null?"":row.getCell(i).toString().toLowerCase());
        }
        //如果未指定name标签，则默认第一列为生成的word文件名
        if (!keys.contains("name")) {
            keys.set(0, "name");
        }
        return keys;
    }

    public static List<Map<String, Object>> extractXls(FileInputStream fis) throws IOException{
        List<Map<String, Object>> results = new ArrayList<Map<String,Object>>();
        try {
            Workbook book = new HSSFWorkbook(new BufferedInputStream(fis));
            Sheet sheet = book.getSheetAt(0);
            if (sheet.getLastRowNum() == 0) {
                return null;
            }
            List<String> keys = getKeyList(sheet.getRow(0));
            for (int i=sheet.getFirstRowNum()+1; i <= sheet.getLastRowNum(); i++) {
                Map<String, Object> param = extractRow(sheet.getRow(i), keys);
                results.add(param);
            }
        } catch (IOException e) {
            throw e;
        }
        return results;
    }

}
