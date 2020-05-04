package wang.crick.excel2word.config.enumerations;

public enum FileType {
    XML_TYPE(".xml"),
    WORD_03_TYPE(".doc"),
    EXCEL_03_TYPE(".xls")
    ;
    private String label;

    private FileType(String label) {
        this.label = label;
    }

    public String getLabel() {
        return label;
    }
}