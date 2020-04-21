package tiho;

import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JPanel;
import javax.swing.JScrollPane;

public class ParameterChooserPanel extends JPanel
{
	private static final long serialVersionUID = 1L;

	private String[] parameters;

	private JButton selectAllButton;
	private JPanel parameterPanel;

	private JCheckBox[] boxes;

	public ParameterChooserPanel(String[] parameters)
	{
		this.parameters = parameters;

		parameterPanel = new JPanel();

		selectAllButton = new JButton("alle auswählen");

		boxes = new JCheckBox[parameters.length];

		layOutComponents();

		selectAllButton.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
				boolean[] areSelected = new boolean[boxes.length];
				for(int i = 0; i < areSelected.length; i++)
				{
					areSelected[i] = boxes[i].isSelected();
				}

				boolean areAllSelected = isAllTrue(areSelected);

				if(!areAllSelected)
				{
					for(JCheckBox box : boxes)
					{
						box.setSelected(true);
					}
				} else
				{
					for(JCheckBox box: boxes)
					{
						box.setSelected(false);
					}
				}
			}
		});
	}

	private boolean isAllTrue(boolean[] bool)
	{
		for(boolean boolv : bool)
		{
			if(!boolv)
			{
				return false;
			}
		}

		return true;
	}

	private void setPanel()
	{
		parameterPanel.setLayout(new GridBagLayout());

		GridBagConstraints gc = new GridBagConstraints();
		gc.weightx = 1;
		gc.weighty = 1;
		gc.fill = GridBagConstraints.NONE;
		gc.anchor = GridBagConstraints.FIRST_LINE_START;
		gc.gridx = 0;
		gc.gridy = 0;

		int counter = 0;

		for(int i = 0; i < parameters.length; i++)
		{
			gc.gridx = counter;
			boxes[i] = new JCheckBox(parameters[i]);
			parameterPanel.add(boxes[i], gc);
			if(counter == 0)
				counter = 1;
			else
			{
				gc.gridy++;
				counter = 0;
			}
		}
	}

	private void layOutComponents()
	{
		setLayout(new GridBagLayout());

		GridBagConstraints gc = new GridBagConstraints();

		gc.weightx = 1;
		gc.weighty = 0.1;

		gc.fill = GridBagConstraints.NONE;

		// FIRST LINE
		gc.gridy = 0;
		gc.gridx = 0;
		gc.anchor = GridBagConstraints.LINE_END;
		gc.insets = new Insets(0, 0, 0, 10);
		add(selectAllButton, gc);

		setPanel();

		gc.insets = new Insets(0, 0, 0, 0);

		gc.fill = GridBagConstraints.BOTH;
		gc.anchor = GridBagConstraints.CENTER;

		parameterPanel.setBorder(BorderFactory.createEmptyBorder(0, 10, 0, 10));

		JPanel parPanel = new JPanel();
		parPanel.setBorder(BorderFactory.createEtchedBorder());
		parPanel.setLayout(new GridBagLayout());

		// NEXT LINE
		gc.gridy++;
		gc.gridx = 0;
		gc.weighty = 1;
		gc.anchor = GridBagConstraints.FIRST_LINE_START;
		JScrollPane scrollPane = new JScrollPane(parameterPanel);
		scrollPane.setBorder(null);
		scrollPane.getVerticalScrollBar().setUnitIncrement(4);
		parPanel.add(scrollPane, gc);

		// NEXT LINE
		JPanel panel = new JPanel();
		gc.gridy++;
		gc.weighty = (8 - Math.floor(parameters.length / 2));
		parPanel.add(panel, gc);

		gc.weighty = 1;
		add(parPanel, gc);
	}
	
	public JCheckBox[] getCheckBoxes()
	{
		return boxes;
	}
}