package excelgrep.gui;

public enum LaunchMode {
    AwtDesktop("AwtDesktop"),
    VBScript("VBScript"), 
    ;
    
    final public static String PROPERTY_KEY = "launch_mode";

    String label;
    LaunchMode(String label) {
        this.label =label;
    }
    
    public static LaunchMode getDefault() {
        return VBScript;
    }
    
    @Override
    public String toString() {
        return label;
    }

}
