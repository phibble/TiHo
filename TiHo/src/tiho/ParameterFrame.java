package tiho;

import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.ArrayList;
import java.util.List;

import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.border.Border;

public class ParameterFrame extends JFrame
{
	private static final long serialVersionUID = 1L;

	private String[] parameters;
	private ParameterChooserPanel parameterChooser;

	private List<String> chosenParametersList;
	private String[] chosenParameters;

	private JPanel mainPanel;
	private JButton confirmButton;
	private JButton backButton;

	@SuppressWarnings("unused")
	private ExcelReader exReader;

	public ParameterFrame(String path, String[] parameters)
	{
		this.parameters = parameters;
		parameterChooser = new ParameterChooserPanel(this.parameters);
		confirmButton = new JButton("OK");
		backButton = new JButton("Zurück");

		confirmButton.setSize(new Dimension(backButton.getPreferredSize().width + 50, backButton.getPreferredSize().height));

		mainPanel = new JPanel();
		Border innerBorder = BorderFactory.createEtchedBorder();
		Border outerBorder = BorderFactory.createEmptyBorder(1, 1, 1, 1);
		mainPanel.setBorder(BorderFactory.createCompoundBorder(outerBorder, innerBorder));

		backButton.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
				ParameterFrame.this.dispose();
				new MainFrame();
			}
		});

		confirmButton.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
				retrieveChosenParameters();

				String par = "";

				for(int i = 0; i < chosenParameters.length; i++)
				{
					par += chosenParameters[i] + (i != chosenParameters.length - 1 ? ", " : "");
				}

				String[] options = new String[] { "Ja", "Nein" };
				int response = JOptionPane.showOptionDialog(ParameterFrame.this,
						"Wollen Sie das Programm mit den gewählten Parametern (" + par + ") starten?", "Start",
						JOptionPane.DEFAULT_OPTION, JOptionPane.PLAIN_MESSAGE, null, options, options[1]);

				if(response == 0)
				{
					exReader = new ExcelReader(path, chosenParameters);
				} else if(response == 1)
				{
					ParameterFrame.this.dispose();
					new ParameterFrame(path, ParameterFrame.this.parameters);
				}
			}
		});

		setFrame();
		setComponentLayout();
	}

	private void retrieveChosenParameters()
	{
		JCheckBox[] boxes = parameterChooser.getCheckBoxes();

		chosenParametersList = new ArrayList<String>();

		for(int i = 0; i < boxes.length; i++)
		{
			if(boxes[i].isSelected())
			{
				chosenParametersList.add(boxes[i].getText());
			}
		}

		chosenParameters = new String[chosenParametersList.size()];

		for(int i = 0; i < chosenParameters.length; i++)
		{
			chosenParameters[i] = chosenParametersList.get(i);
		}
	}

	private void setComponentLayout()
	{
		mainPanel.setLayout(new GridBagLayout());

		GridBagConstraints gc = new GridBagConstraints();

		gc.weightx = 1;
		gc.weighty = 1;

		gc.fill = GridBagConstraints.NONE;

		// FIRST LINE
		gc.gridy = 0;

		gc.fill = GridBagConstraints.BOTH;

		gc.gridx = 0;
		gc.gridwidth = 2;
		gc.anchor = GridBagConstraints.CENTER;
		mainPanel.add(parameterChooser, gc);

		// NEXT LINE
		gc.gridy++;

		gc.weighty = 0.1;

		gc.fill = GridBagConstraints.NONE;

		gc.gridwidth = 1;

		gc.gridx = 1;
		gc.anchor = GridBagConstraints.LINE_END;
		gc.insets = new Insets(0, 15, 0, 15);
		mainPanel.add(confirmButton, gc);

		gc.gridx = 0;
		gc.anchor = GridBagConstraints.LINE_START;
		mainPanel.add(backButton, gc);

		setLayout(new BorderLayout());
		add(mainPanel, BorderLayout.CENTER);
	}

	private void setFrame()
	{
		setSize(230, 300);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setVisible(true);
		setResizable(true);
		setLocationRelativeTo(null);
	}
}