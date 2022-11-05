package excelgrep.gui;

import java.awt.BorderLayout;
import java.awt.EventQueue;
import java.awt.Frame;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import javax.swing.JComboBox;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTabbedPane;
import javax.swing.border.EmptyBorder;
import javax.swing.JButton;
import java.awt.FlowLayout;

public class PreferenceDialog extends JDialog {

    ExcelGrepGuiMain guiMain;
    private JComboBox<LaunchMode> launchModeComboBox;
    private JComboBox<GrepMode> grepModeComboBox;
    private JButton okButton;
    private JButton cancelButton;

    /**
     * Create the frame.
     */
    public PreferenceDialog(Frame owner, ExcelGrepGuiMain guiMain) {
        super(owner,"設定", true);
        
        this.guiMain = guiMain;
        
        Configuration configration = guiMain.getConfigration();
        
        setBounds(100, 100, 450, 300);
        JPanel contentPane;
        contentPane = new JPanel();
        contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
        contentPane.setLayout(new BorderLayout(0, 0));
        setContentPane(contentPane);
        
        JTabbedPane tabbedPane = new JTabbedPane(JTabbedPane.LEFT);
        contentPane.add(tabbedPane, BorderLayout.CENTER);
        
        JPanel panel = new JPanel();
        tabbedPane.addTab("基本", null, panel, null);
        GridBagLayout gbl_panel = new GridBagLayout();
        gbl_panel.columnWidths = new int[]{0, 0, 0};
        gbl_panel.rowHeights = new int[]{0, 0, 0};
        gbl_panel.columnWeights = new double[]{0.0, 1.0, Double.MIN_VALUE};
        gbl_panel.rowWeights = new double[]{0.0, 0.0, Double.MIN_VALUE};
        panel.setLayout(gbl_panel);
        
        JLabel lblNewLabel = new JLabel("起動方法");
        GridBagConstraints gbc_lblNewLabel = new GridBagConstraints();
        gbc_lblNewLabel.insets = new Insets(0, 0, 5, 5);
        gbc_lblNewLabel.gridx = 0;
        gbc_lblNewLabel.gridy = 0;
        panel.add(lblNewLabel, gbc_lblNewLabel);
        
        launchModeComboBox = new JComboBox<LaunchMode>();
        launchModeComboBox.addItem(LaunchMode.AwtDesktop);
        launchModeComboBox.addItem(LaunchMode.VBScript);
        GridBagConstraints gbc_launchModeComboBox = new GridBagConstraints();
        gbc_launchModeComboBox.insets = new Insets(0, 0, 5, 0);
        gbc_launchModeComboBox.fill = GridBagConstraints.HORIZONTAL;
        gbc_launchModeComboBox.gridx = 1;
        gbc_launchModeComboBox.gridy = 0;
        panel.add(launchModeComboBox, gbc_launchModeComboBox);
        
        JPanel panel_1 = new JPanel();
        FlowLayout flowLayout = (FlowLayout) panel_1.getLayout();
        flowLayout.setAlignment(FlowLayout.TRAILING);
        contentPane.add(panel_1, BorderLayout.SOUTH);
        
        JLabel lblKennsa = new JLabel("検索方法");
        GridBagConstraints gbc_lblKennsa = new GridBagConstraints();
        gbc_lblKennsa.anchor = GridBagConstraints.EAST;
        gbc_lblKennsa.insets = new Insets(0, 0, 0, 5);
        gbc_lblKennsa.gridx = 0;
        gbc_lblKennsa.gridy = 1;
        panel.add(lblKennsa, gbc_lblKennsa);
        
        grepModeComboBox = new JComboBox<GrepMode>();
        grepModeComboBox.addItem(GrepMode.MultiThread);
        grepModeComboBox.addItem(GrepMode.SingleThread);

        GridBagConstraints gbc_grepModeComboBox = new GridBagConstraints();
        gbc_grepModeComboBox.fill = GridBagConstraints.HORIZONTAL;
        gbc_grepModeComboBox.gridx = 1;
        gbc_grepModeComboBox.gridy = 1;
        panel.add(grepModeComboBox, gbc_grepModeComboBox);

        launchModeComboBox.setSelectedItem(configration.getLaunchMode());
        grepModeComboBox.setSelectedItem(configration.getGrepMode());
        
        okButton = new JButton("OK");
        panel_1.add(okButton);
        okButton.addActionListener( (e)->{
            onClick_okButton(guiMain);
        });
        
        cancelButton = new JButton("キャンセル");
        panel_1.add(cancelButton);
        cancelButton.addActionListener( (e)->{
            PreferenceDialog.this.setVisible(false);
        });
        
    }

    protected void onClick_okButton(ExcelGrepGuiMain guiMain) {
        Configuration configuration = new Configuration();
        configuration.setLaunchMode((LaunchMode) launchModeComboBox.getSelectedItem());
        configuration.setGrepMode((GrepMode) grepModeComboBox.getSelectedItem());
        guiMain.setConfiguration(configuration);
        PreferenceDialog.this.setVisible(false);
    }

}
