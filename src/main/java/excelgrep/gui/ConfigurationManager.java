package excelgrep.gui;

import java.io.File;
import java.io.FileInputStream;
import java.nio.file.Path;
import java.util.Comparator;
import java.util.Properties;
import java.util.Set;
import java.util.TreeSet;
import java.util.Map.Entry;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public class ConfigurationManager {
    static Logger log = LogManager.getLogger(ConfigurationManager.class);
    
    public Configuration loadFile(File filepath) {
        Configuration configuration = new Configuration();
        
        try {
            Properties prop = new Properties();
            prop.load(new FileInputStream(filepath));

            Comparator<Entry<Object, Object>> comparing = Comparator.comparing((it)->it.getKey().toString(), Comparator.naturalOrder());
            Set<Entry<Object, Object>> entrySet = new TreeSet<>(comparing);
            entrySet.addAll(prop.entrySet());
            
            for (Entry<Object, Object> p : entrySet) {
                String key = p.getKey().toString();
                String value = p.getValue().toString();
                switch(key) {
                    case LaunchMode.PROPERTY_KEY:
                        configuration.setLaunchMode(LaunchMode.valueOf(value));
                        break;
                    case GrepMode.PROPERTY_KEY:
                        configuration.setGrepMode(GrepMode.valueOf(value));
                        break;
                }
            }
        } catch (Exception e) {
            log.warn("failed to load properties.", e);
        }
        
        return configuration ;
        
    }

    public void saveFile(File file) {
        
    }
}
