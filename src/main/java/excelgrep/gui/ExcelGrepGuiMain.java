package excelgrep.gui;

import java.awt.BorderLayout;
import java.awt.Desktop;
import java.awt.EventQueue;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.GridLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.Comparator;
import java.util.Map.Entry;
import java.util.Properties;
import java.util.Set;
import java.util.TreeSet;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.ListSelectionModel;
import javax.swing.SwingConstants;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import excelgrep.core.data.ExcelData;

public class ExcelGrepGuiMain {
    private static final String SETTING_PREFIX_KEYWORD = "keyword";

    private static final String SETTING_PREFIX_FOLDER = "folder";

    private static final String SETTING_PROPERTIES = "excelgrep.properties";
    private static final String CONFIGURATION_PROPERTIES = "config.properties";

    static Logger log = LogManager.getLogger(ExcelGrepGuiMain.class);

    // 
    private Configuration configuration;
    private Properties prop;
    
    // GUIs
    private JFrame frame;
    private JTable table;
    private JPanel panel;
    private JPanel operationPanel;
    private JPanel panel_1;
    private JLabel lblNewLabel;
    private JComboBox<String> searchFolder;
    private JButton selectFolderButton;
    private JLabel lblNewLabel_1;
    private JComboBox<String> searchKeyword;
    private JButton searchButton;
    private JPanel panel_2;
    JLabel statusBar;
    JProgressBar progressBar;
    private JProgressBar progressBar_1;
    private ExcelGrepSearchWorker searchWorker;

    private JMenuBar menuBar;
    private JMenu menuCategoryFile;
    private JMenuItem menuSettings;

    /**
     * Launch the application.
     * 
     * @throws UnsupportedLookAndFeelException
     * @throws IllegalAccessException
     * @throws InstantiationException
     * @throws ClassNotFoundException
     */
    public static void main(String[] args) throws ClassNotFoundException, InstantiationException, IllegalAccessException, UnsupportedLookAndFeelException {
        UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());

        EventQueue.invokeLater(new Runnable() {
            public void run() {
                try {
                    ExcelGrepGuiMain window = new ExcelGrepGuiMain();
                    window.frame.setVisible(true);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        });
    }

    /**
     * Create the application.
     */
    public ExcelGrepGuiMain() {
        initialize();
        loadConfiguration();
        loadProperty();
    }

    private void loadConfiguration() {
        ConfigurationManager configurationManager = new ConfigurationManager();
        this.configuration = configurationManager.loadFile(new File(CONFIGURATION_PROPERTIES));
    }

    public void setConfiguration(Configuration configuration) {
        ConfigurationManager configurationManager = new ConfigurationManager();
        configurationManager.saveFile(new File(CONFIGURATION_PROPERTIES), configuration);
        this.configuration = configuration;
    }
    
    public Configuration getConfigration() {
        return this.configuration;
        
    }

    private void loadProperty() {
        try {
            prop = new Properties();
            prop.load(new FileInputStream(new File(SETTING_PROPERTIES)));

            Comparator<Entry<Object, Object>> comparing = Comparator.comparing((it)->it.getKey().toString(), Comparator.naturalOrder());
            Set<Entry<Object, Object>> entrySet = new TreeSet<>(comparing);
            entrySet.addAll(prop.entrySet());
            
            for (Entry<Object, Object> p : entrySet) {
                String key = p.getKey().toString();
                if (key.startsWith(SETTING_PREFIX_FOLDER)) {
                    searchFolder.addItem(p.getValue().toString());
                } else if (key.startsWith(SETTING_PREFIX_KEYWORD)) {
                    searchKeyword.addItem(p.getValue().toString());
                }
            }
        } catch (Exception e) {
            log.warn("failed to load properties.", e);
            throw new RuntimeException(e);
        }
    }

    private void saveProperties() {
        try {
            prop = new Properties();
            for (int i = 0; i < searchFolder.getItemCount(); i++) {
                prop.setProperty(SETTING_PREFIX_FOLDER + i, searchFolder.getItemAt(i));
            }
            for (int i = 0; i < searchKeyword.getItemCount(); i++) {
                prop.setProperty(SETTING_PREFIX_KEYWORD + i, searchKeyword.getItemAt(i));
            }
            prop.store(new FileOutputStream(SETTING_PROPERTIES), "");
        } catch (IOException e) {
            log.warn("failed to save properties.", e);
            throw new RuntimeException(e);
        }
    }

    /**
     * Initialize the contents of the frame.
     */
    private void initialize() {
        frame = new JFrame();
        frame.setBounds(100, 100, 450, 300);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        ExcelGrepResultTableModel dataModel = new ExcelGrepResultTableModel(new String[] {"\u30D1\u30B9", "\u30B7\u30FC\u30C8", "\u5834\u6240", "\u30C6\u30AD\u30B9\u30C8"});

        table = new JTable();
        table.setEnabled(false);
        table.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
        table.setFillsViewportHeight(true);
        table.setAutoCreateRowSorter(true);
        table.setModel(dataModel);
        table.getColumnModel().getColumn(0).setPreferredWidth(237);
        table.getColumnModel().getColumn(1).setPreferredWidth(38);
        table.getColumnModel().getColumn(2).setPreferredWidth(49);
        table.getColumnModel().getColumn(3).setPreferredWidth(253);
        JScrollPane mainPanel = new JScrollPane(table);

        frame.getContentPane().add(mainPanel, BorderLayout.CENTER);

        operationPanel = new JPanel();
        frame.getContentPane().add(operationPanel, BorderLayout.NORTH);
        operationPanel.setLayout(new GridLayout(0, 1, 0, 0));

        panel_1 = new JPanel();
        operationPanel.add(panel_1);
        GridBagLayout gbl_panel_1 = new GridBagLayout();
        gbl_panel_1.columnWidths = new int[] {50, 139, 91, 0};
        gbl_panel_1.rowHeights = new int[] {21, 0, 0};
        gbl_panel_1.columnWeights = new double[] {0.0, 1.0, 0.0, Double.MIN_VALUE};
        gbl_panel_1.rowWeights = new double[] {0.0, 0.0, Double.MIN_VALUE};
        panel_1.setLayout(gbl_panel_1);

        lblNewLabel = new JLabel("検索ディレクトリ");
        lblNewLabel.setHorizontalAlignment(SwingConstants.LEFT);
        GridBagConstraints gbc_lblNewLabel = new GridBagConstraints();
        gbc_lblNewLabel.insets = new Insets(0, 5, 0, 5);
        gbc_lblNewLabel.anchor = GridBagConstraints.WEST;
        gbc_lblNewLabel.gridx = 0;
        gbc_lblNewLabel.gridy = 0;
        panel_1.add(lblNewLabel, gbc_lblNewLabel);

        searchFolder = new JComboBox<>();
        searchFolder.setEditable(true);
        GridBagConstraints gbc_searchFolder = new GridBagConstraints();
        gbc_searchFolder.insets = new Insets(0, 0, 5, 5);
        gbc_searchFolder.anchor = GridBagConstraints.NORTH;
        gbc_searchFolder.fill = GridBagConstraints.HORIZONTAL;
        gbc_searchFolder.gridx = 1;
        gbc_searchFolder.gridy = 0;
        panel_1.add(searchFolder, gbc_searchFolder);

        selectFolderButton = new JButton("選択");
        GridBagConstraints gbc_selectFolderButton = new GridBagConstraints();
        gbc_selectFolderButton.insets = new Insets(0, 5, 5, 0);
        gbc_selectFolderButton.anchor = GridBagConstraints.NORTHEAST;
        gbc_selectFolderButton.gridx = 2;
        gbc_selectFolderButton.gridy = 0;
        panel_1.add(selectFolderButton, gbc_selectFolderButton);

        lblNewLabel_1 = new JLabel("キーワード");
        lblNewLabel_1.setHorizontalAlignment(SwingConstants.LEFT);
        GridBagConstraints gbc_lblNewLabel_1 = new GridBagConstraints();
        gbc_lblNewLabel_1.anchor = GridBagConstraints.WEST;
        gbc_lblNewLabel_1.insets = new Insets(0, 5, 0, 5);
        gbc_lblNewLabel_1.gridx = 0;
        gbc_lblNewLabel_1.gridy = 1;
        panel_1.add(lblNewLabel_1, gbc_lblNewLabel_1);

        searchKeyword = new JComboBox<>();
        searchKeyword.setEditable(true);
        GridBagConstraints gbc_searchKeyword = new GridBagConstraints();
        gbc_searchKeyword.insets = new Insets(0, 0, 0, 5);
        gbc_searchKeyword.fill = GridBagConstraints.HORIZONTAL;
        gbc_searchKeyword.gridx = 1;
        gbc_searchKeyword.gridy = 1;
        panel_1.add(searchKeyword, gbc_searchKeyword);

        searchButton = new JButton("検索");
        GridBagConstraints gbc_searchButton = new GridBagConstraints();
        gbc_searchButton.anchor = GridBagConstraints.EAST;
        gbc_searchButton.gridx = 2;
        gbc_searchButton.gridy = 1;
        panel_1.add(searchButton, gbc_searchButton);

        panel_2 = new JPanel();
        frame.getContentPane().add(panel_2, BorderLayout.SOUTH);
        GridBagLayout gbl_panel_2 = new GridBagLayout();
        gbl_panel_2.columnWidths = new int[] {0, 0, 0, 0};
        gbl_panel_2.rowHeights = new int[] {0, 0};
        gbl_panel_2.columnWeights = new double[] {0.0, 0.0, 0.0, Double.MIN_VALUE};
        gbl_panel_2.rowWeights = new double[] {0.0, Double.MIN_VALUE};
        panel_2.setLayout(gbl_panel_2);

        statusBar = new JLabel("...");
        GridBagConstraints gbc_statusBar = new GridBagConstraints();
        gbc_statusBar.weightx = 1.0;
        gbc_statusBar.fill = GridBagConstraints.BOTH;
        gbc_statusBar.insets = new Insets(0, 0, 5, 5);
        gbc_statusBar.gridx = 0;
        gbc_statusBar.gridy = 0;
        panel_2.add(statusBar, gbc_statusBar);

        progressBar = new JProgressBar();
        GridBagConstraints gbc_progressBar = new GridBagConstraints();
        gbc_progressBar.gridx = 1;
        gbc_progressBar.gridy = 0;
        panel_2.add(progressBar, gbc_progressBar);

        progressBar_1 = new JProgressBar();
        GridBagConstraints gbc_progressBar_1 = new GridBagConstraints();
        gbc_progressBar_1.anchor = GridBagConstraints.EAST;
        gbc_progressBar_1.gridx = 2;
        gbc_progressBar_1.gridy = 0;
        panel_2.add(progressBar_1, gbc_progressBar_1);
        
        menuBar = new JMenuBar();
        frame.setJMenuBar(menuBar);
        
        menuCategoryFile = new JMenu("ファイル");
        menuBar.add(menuCategoryFile);
        
        menuSettings = new JMenuItem("設定");
        menuSettings.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                PreferenceDialog dialog = new PreferenceDialog(frame, ExcelGrepGuiMain.this);
                dialog.setVisible(true);
            }
        });
        menuCategoryFile.add(menuSettings);

        registerEventHandlers();
    }

    private void registerEventHandlers() {
        selectFolderButton.addActionListener(this::onClick_selectFolderButton);
        searchButton.addActionListener(this::onClick_searchButton);
        table.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                launchExcel(e);
            }


        });

    }

    private void launchExcel(MouseEvent e) {
        if (e.getClickCount() != 2) {
            return;
        }
        int idx = table.getSelectedRow();
        if (idx == -1) {
            return;
        }
        idx = table.convertRowIndexToModel(idx);
        ExcelGrepResultTableModel model = (ExcelGrepResultTableModel) table.getModel();

        ExcelData row = model.getRow(idx);

        switch( configuration.getLaunchMode() ) {
            case Default:
                launchExcelUsingDesktop(row);
                break;
            case VBScript:
                launchExcelUsingVbs(row);
                break;
        }
    }

    private void launchExcelUsingVbs(ExcelData row) {
        try {
            Path filepath = (Path) row.getPosition().getFilePath();
            String sheet = (String) row.getPosition().getSheetName();
            String cell = (String) row.getPosition().getCellPosition();
            
            // wscriptへ引数を渡す際、      でエスケープの問題を回避。
            // 
            String encodedFilepath = encode(filepath.toString());
            String encodedSheet =  encode(sheet.toString());
            String encodedCell =  encode(cell.toString());
            ProcessBuilder builder = new ProcessBuilder("wscript", "launchExcel.vbs", encodedFilepath, encodedSheet, encodedCell );
            builder.start();
         }catch (Exception e) {
             throw new RuntimeException(e);
         }
    }
    private String encode(String value) throws Exception {
        return URLEncoder.encode(value,  "MS932");
    }

    private void launchExcelUsingDesktop(ExcelData row) {
        if (!Desktop.isDesktopSupported()) {
            return;
        }
        try {
            Path filepath = (Path) row.getPosition().getFilePath();
            String sheet = (String) row.getPosition().getSheetName();
            
            Desktop.getDesktop().open(filepath.toFile());
        } catch (IOException ex) {
            log.error("failed to launch", ex);
        }
    }

    void onClick_selectFolderButton(ActionEvent e) {
        String oldFolderPath = (String) searchFolder.getSelectedItem();

        JFileChooser filechooser = new JFileChooser();
        filechooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        if (oldFolderPath != null) {
            filechooser.setCurrentDirectory(new File(oldFolderPath));
        }

        int selected = filechooser.showOpenDialog(frame);
        if (selected == JFileChooser.APPROVE_OPTION) {
            File selectedFile = filechooser.getSelectedFile();
            String newFolderPath = selectedFile.getAbsolutePath();
            searchFolder.setSelectedItem(newFolderPath);
        }
    }

    void onClick_searchButton(ActionEvent e) {
        String targetFolderPath = (String) searchFolder.getSelectedItem();
        String regex = (String) searchKeyword.getSelectedItem();

        if (targetFolderPath == null) {
            return;
        }

        if (regex == null) {
            return;
        }
        
        searchFolder.removeItem(targetFolderPath);
        searchKeyword.removeItem(regex);
        
        searchFolder.insertItemAt(targetFolderPath , 0 );
        searchKeyword.insertItemAt(regex, 0);

        searchFolder.setSelectedItem(targetFolderPath);
        searchKeyword.setSelectedItem(regex);
        
        saveProperties();

        // searchWorker = new ExcelGrepSearchWorker(this, targetFolderPath, regex, table);
        searchWorker = new ExcelGrepSearchMultiThreadWorker(this, targetFolderPath, regex, table);
        
        searchWorker.execute();

    }


    public void updateStatusBar(String text) {
        statusBar.setText(text);
    }


}
