package tiho;

import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.Arrays;
import java.util.List;

import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.UIManager;
import javax.swing.border.Border;

public class MainFrame extends JFrame
{
	private static final long serialVersionUID = 1L;

	private ChoosePanel choosePanel;
	private JPanel framePanel;
	private JButton confirmButton;
	private String fileAbsolutePath = "";
	private boolean allParameters = true;

	private ParameterFinder paramFinder;
	private FileSaver fileSaver;

	private static List<String> prevFiles;

	private GridBagConstraints gc;

	@SuppressWarnings("unused")
	private ExcelReader exReader;

	@SuppressWarnings("unused")
	private ParameterFrame parameterFrame;

	public MainFrame()
	{
		setProgramLookAndFeel();

		String[] prev = FileSaver.readPrevFile().split(",");
		prevFiles = Arrays.asList(prev);

		choosePanel = new ChoosePanel();
		framePanel = new JPanel();
		confirmButton = new JButton("OK");

		paramFinder = null;
		fileSaver = null;

		setFrameSettings();
		setComponentLayout();

		confirmButton.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
				fileAbsolutePath = choosePanel.getFileAbsolutePath();

				fileSaver = new FileSaver(fileAbsolutePath);

				if(!(prevFiles.contains(fileAbsolutePath)))
				{
					fileSaver.writePrevFile();
				}
				allParameters = choosePanel.isAllParameters();

				if(fileAbsolutePath != "")
				{
					if(!allParameters)
					{
						paramFinder = new ParameterFinder(fileAbsolutePath);
						setupParameterPanel();
					} else
					{
						String[] options = new String[] { "Ja", "Nein" };
						int response = JOptionPane.showOptionDialog(MainFrame.this,
								"Wollen Sie das Programm mit allen Parametern starten?", "Start",
								JOptionPane.DEFAULT_OPTION, JOptionPane.PLAIN_MESSAGE, null, options, options[1]);

						if(response == 0)
						{
							exReader = new ExcelReader(fileAbsolutePath);
						} else if(response == 1)
						{
							MainFrame.this.dispose();
							new MainFrame();
						}
					}
				} else
				{
					MainFrame.this.dispose();
					new MainFrame();
					JOptionPane.showMessageDialog(MainFrame.this, "Es ist keine Datei ausgewählt!");
				}
			}
		});
	}

	private void setupParameterPanel()
	{
		this.dispose();

		this.parameterFrame = new ParameterFrame(fileAbsolutePath, paramFinder.getParameters());
	}

	private void setProgramLookAndFeel()
	{
		try
		{
			UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
		} catch(Exception e)
		{
			e.printStackTrace();
		}
	}

	private void setFrameSettings()
	{
		setTitle("Dateiauswahl");

		setSize(360, 145);
		setVisible(true);
		setResizable(false);
		setLocationRelativeTo(null);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

		Border innerBorder = BorderFactory.createEtchedBorder();
		Border outerBorder = BorderFactory.createEmptyBorder(1, 1, 1, 1);
		framePanel.setBorder(BorderFactory.createCompoundBorder(outerBorder, innerBorder));
		setPreferredSize(new Dimension(305, 195));
	}

	private void setComponentLayout()
	{
		setLayout(new BorderLayout());
		framePanel.setLayout(new GridBagLayout());

		gc = new GridBagConstraints();

		gc.weightx = 1;
		gc.weighty = 1;

		gc.fill = GridBagConstraints.NONE;

		gc.gridx = 0;
		gc.gridy = 0;
		framePanel.add(choosePanel, gc);

		gc.gridx = 0;
		gc.gridy = 1;
		framePanel.add(confirmButton, gc);
		add(framePanel, BorderLayout.CENTER);
	}

	public String getFileAbsolutePath()
	{
		return fileAbsolutePath;
	}

	public static List<String> getPrevFiles()
	{
		return prevFiles;
	}
}