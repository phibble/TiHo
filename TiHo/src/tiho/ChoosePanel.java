package tiho;

import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.List;

import javax.swing.BorderFactory;
import javax.swing.ButtonGroup;
import javax.swing.DefaultComboBoxModel;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JRadioButton;

public class ChoosePanel extends JPanel
{
	private static final long serialVersionUID = 1L;

	private JComboBox<String> fileName;
	private DefaultComboBoxModel<String> fileNameModel;
	private JLabel fileNameLabel;
	private JButton searchButton;
	private JFileChooser fileChooser;

	private JRadioButton allRadio;
	private JRadioButton someRadio;
	private ButtonGroup allRadioGroup;
	private JPanel radioPanel;

	private String fileAbsolutePath = "";

	private boolean allParameters = true;

	public ChoosePanel()
	{
		fileName = new JComboBox<String>();
		fileNameLabel = new JLabel("Datei: ");
		searchButton = new JButton("...");
		fileChooser = new JFileChooser();

		fileChooser.setFileFilter(new XLSXFilter());

		allRadio = new JRadioButton("mit allen Parametern starten");
		someRadio = new JRadioButton("Parameter wählen");
		radioPanel = new JPanel();
		allRadioGroup = new ButtonGroup();
		allRadioGroup.add(allRadio);
		allRadioGroup.add(someRadio);
		allRadio.setSelected(true);

		radioPanel.setLayout(new BorderLayout());
		radioPanel.add(allRadio, BorderLayout.NORTH);
		radioPanel.add(someRadio, BorderLayout.SOUTH);
		radioPanel.setBorder(BorderFactory.createEmptyBorder(5, 10, 5, 0));

		fileNameModel = new DefaultComboBoxModel<>();
		fileNameModel.addElement("");

		List<String> prevFiles = MainFrame.getPrevFiles();

		for(int i = 0; i < prevFiles.size(); i++)
		{
			fileNameModel.addElement(prevFiles.get(i));
		}

		fileName.setModel(fileNameModel);
		fileName.setSelectedIndex(prevFiles.size());
		fileName.setEditable(true);
		fileName.setPreferredSize(new Dimension(270, 20));

		searchButton.setPreferredSize(new Dimension(22, 19));

		searchButton.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
				if(fileChooser.showOpenDialog(ChoosePanel.this) == JFileChooser.APPROVE_OPTION)
				{
					fileAbsolutePath = fileChooser.getSelectedFile().getAbsolutePath();

					if(!prevFiles.contains(fileAbsolutePath))
					{
						fileNameModel.addElement(fileAbsolutePath);
					}

					FileSaver saver = new FileSaver(fileAbsolutePath);
					if(fileNameModel.getIndexOf(fileAbsolutePath) != fileNameModel.getSize() - 1)
					{
						saver.changeLastFile();
					}
					fileNameModel.setSelectedItem(fileAbsolutePath);
				}
			}
		});

		allRadio.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
				allParameters = true;
			}
		});

		someRadio.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
				allParameters = false;
			}
		});

		setComponentLayout();
	}

	private void setComponentLayout()
	{
		setLayout(new GridBagLayout());

		GridBagConstraints gc = new GridBagConstraints();

		gc.weightx = 1;
		gc.weighty = 1;

		gc.fill = GridBagConstraints.NONE;

		// FIRST ROW
		gc.gridy = 0;

		gc.gridx = 1;
		gc.anchor = GridBagConstraints.LINE_START;
		add(fileName, gc);

		gc.gridx = 0;
		gc.anchor = GridBagConstraints.LINE_START;
		add(fileNameLabel, gc);

		gc.gridx = 2;
		gc.anchor = GridBagConstraints.CENTER;
		gc.weightx = 2;
		add(searchButton, gc);

		// NEXT ROW
		gc.gridy++;
		gc.gridx = 0;
		gc.gridwidth = 3;
		gc.gridheight = 2;
		gc.weighty = 1;
		gc.anchor = GridBagConstraints.LINE_START;
		add(radioPanel, gc);
	}

	public String getFileAbsolutePath()
	{
		return (String) fileName.getSelectedItem();
	}

	public boolean isAllParameters()
	{
		return allParameters;
	}
}