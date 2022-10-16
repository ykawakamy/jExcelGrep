package excelgrep.gui;

public enum GrepMode {
    SingleThread("デフォルト"),
    MultiThread("VBS"),
    ;

    final public static String PROPERTY_KEY = "grep_mode";

    String label;
    GrepMode(String label) {
        this.label =label;
    }
    
    public static GrepMode getDefault() {
        return SingleThread;
    }
    
}
