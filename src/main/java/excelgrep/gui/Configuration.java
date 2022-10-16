package excelgrep.gui;

public class Configuration {
    private LaunchMode launchMode = LaunchMode.getDefault();
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
