package excelgrep.gui;

public enum LaunchMode {
    Default("デフォルト"),
    VBScript("VBS"), 
    ;
    
    final public static String PROPERTY_KEY = "launch_mode";

    String label;
    LaunchMode(String label) {
        this.label =label;
    }
    
    public static LaunchMode getDefault() {
        return Default;
    }
    
}
