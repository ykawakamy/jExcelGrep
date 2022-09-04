package excelgrep.gui;

import java.awt.BorderLayout;
import java.awt.EventQueue;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.GridLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.io.File;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.SwingConstants;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
import javax.swing.table.DefaultTableModel;

public class ExcelGrepGuiMain {

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
    private JLabel statusBar;
    private JProgressBar progressBar;
    private JProgressBar progressBar_1;
    private ExcelGrepSearchWorker searchWorker;

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
    }

    /**
     * Initialize the contents of the frame.
     */
    private void initialize() {
        frame = new JFrame();
        frame.setBounds(100, 100, 450, 300);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);


        table = new JTable();
        table.setFillsViewportHeight(true);
        table.setModel(new DefaultTableModel(new Object[][] {{null, null, null, null},}, new String[] {"\u30D1\u30B9", "\u30B7\u30FC\u30C8", "\u5834\u6240", "\u30C6\u30AD\u30B9\u30C8"}));
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
        panel_2.add(progressBar_1, gbc_progressBar_1);;

        registerEventHandlers();
    }

    private void registerEventHandlers() {
        selectFolderButton.addActionListener(this::onClick_selectFolderButton);
        searchButton.addActionListener(this::onClick_searchButton);
        
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
            searchFolder.addItem(newFolderPath);
            searchFolder.setSelectedItem(newFolderPath);
        }
    }
    
    void onClick_searchButton(ActionEvent e){
        String targetFolderPath = (String) searchFolder.getSelectedItem();
        String regex = (String)searchKeyword.getSelectedItem();
        
        searchWorker = new ExcelGrepSearchWorker(targetFolderPath , regex);
        
    }

}
