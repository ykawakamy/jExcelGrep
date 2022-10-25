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
        tabbedPane.addTab("New tab", null, panel, null);
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
        
        JComboBox<LaunchMode> launchModeComboBox = new JComboBox<LaunchMode>();
        launchModeComboBox.addItem(LaunchMode.AwtDesktop);
        launchModeComboBox.addItem(LaunchMode.VBScript);
        GridBagConstraints gbc_comboBox = new GridBagConstraints();
        gbc_comboBox.insets = new Insets(0, 0, 5, 0);
        gbc_comboBox.fill = GridBagConstraints.HORIZONTAL;
        gbc_comboBox.gridx = 1;
        gbc_comboBox.gridy = 0;
        panel.add(launchModeComboBox, gbc_comboBox);
        
        JPanel panel_1 = new JPanel();
        FlowLayout flowLayout = (FlowLayout) panel_1.getLayout();
        flowLayout.setAlignment(FlowLayout.TRAILING);
        contentPane.add(panel_1, BorderLayout.SOUTH);
        
        JButton okButton = new JButton("OK");
        panel_1.add(okButton);
        okButton.addActionListener( (e)->{
            
            Configuration configuration = new Configuration();
            configuration.setLaunchMode((LaunchMode) launchModeComboBox.getSelectedItem());
            guiMain.setConfiguration(configuration);
            PreferenceDialog.this.setVisible(false);
        });

        
        JButton cancelButton = new JButton("キャンセル");
        panel_1.add(cancelButton);
        cancelButton.addActionListener( (e)->{
            PreferenceDialog.this.setVisible(false);
        });
        
        launchModeComboBox.setSelectedItem(configration.getLaunchMode());
    }

}
