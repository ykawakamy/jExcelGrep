package excelgrep.gui;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Path;
import java.util.Comparator;
import java.util.Properties;
import java.util.Set;
import java.util.TreeSet;
import java.util.Map.Entry;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

/**
 * コンフィグのファイル書き出し/読み込みを行う。
 */
public class ConfigurationManager {
    static Logger log = LogManager.getLogger(ConfigurationManager.class);
    
    /**
     * 指定したファイルパスからコンフィグを読み込む
     * <p>読み込みに失敗した場合、デフォルトを返す。</p>
     * @param filepath ファイルパス
     * @return 読み込みしたコンフィグ
     */
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
            
            return new Configuration();
        }
        
        return configuration ;
        
    }

    /**
     * 指定したファイルパスにコンフィグを書き出す。
     * @param filepath 書き出し先のファイルパス
     * @param configuration 書き出すコンフィグ
     */
    public void saveFile(File filepath, Configuration configuration) {
        
        try {
            Properties prop = new Properties();

            prop.put(LaunchMode.PROPERTY_KEY,configuration.getLaunchMode().name());
            prop.put(GrepMode.PROPERTY_KEY,configuration.getGrepMode().name());
            
            prop.store(new FileOutputStream(filepath), "");
        } catch (Exception e) {
            log.warn("failed to save properties.", e);
            throw new RuntimeException(e);
        }
    }
}
