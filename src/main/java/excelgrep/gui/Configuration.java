package excelgrep.gui;

/**
 * コンフィグ
 */
public class Configuration {
    
    /** 起動方式 */
    private LaunchMode launchMode = LaunchMode.getDefault();
    /** Grep方式 */
    private GrepMode grepMode = GrepMode.getDefault();
    
    public LaunchMode getLaunchMode() {
        return launchMode;
    }
    public void setLaunchMode(LaunchMode launchMode) {
        this.launchMode = launchMode;
    }
    
    public GrepMode getGrepMode() {
        return grepMode;
    }
    public void setGrepMode(GrepMode grepMode) {
        this.grepMode = grepMode;
    }

}
