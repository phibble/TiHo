package tiho;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import javax.swing.JFrame;
import javax.swing.JOptionPane;

import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader
{
	private String path;

	private XSSFSheet sheet;
	private XSSFWorkbook workbook;
	private FileInputStream file;

	@SuppressWarnings("unused")
	private ExcelWriter exWriter;

	public ExcelReader(String path, String[] parameters)
	{
		this.path = path;

		getExcelSheet();
		removeExcelSheets(workbook);

		readExcelFile(parameters);
	}

	private void readExcelFile(String[] parameters)
	{
		Row row = null;
		Cell cell = null;

		int rowCounter = 0;
		int cellCounter = 0;

		Probe probe = new Probe();
		
		boolean amino = false;
		List<String> aminoProbes = new ArrayList<String>();

		for(int rowIterator = 2; rowIterator <= sheet.getLastRowNum(); rowIterator++)
		{
			row = sheet.getRow(rowIterator);

			if(isRowEmpty(row))
			{
				continue;
			}

			for(int cellIterator = 0; cellIterator <= row.getLastCellNum(); cellIterator++)
			{
				cell = row.getCell(cellIterator);

				if(cell == null)
				{
					continue;
				}

				if(cell.getCellTypeEnum() == CellType.STRING)
				{
					if(cell.getStringCellValue().trim().toLowerCase().contains("probe") && rowCounter != 0)
					{
						rowCounter = 0;
						cellCounter = 0;
						amino = false;
					}
					
					if(rowCounter == 0 && cellCounter == 0)
					{
						probe.setNumber(cell.getStringCellValue());
					} else if(rowCounter == 0 && cellCounter == 1)
					{
						probe.setName(cell.getStringCellValue());
					} else if(cellCounter == 0 && rowCounter != 0)
					{
						writeExcelSheet(row, probe, parameters, amino, aminoProbes);
					}
					
					if(cell.getStringCellValue().trim().toLowerCase().contains("aminosäure"))
					{
						if(!aminoProbes.contains(probe.getName()))
						{
							aminoProbes.add(probe.getName());
						}
						amino = true;
					}
				}
				cellCounter++;
			}

			cellCounter = 0;
			rowCounter++;
		}
		writeSheet(workbook);
		closeFile();
		System.exit(0);
	}

	private void closeFile()
	{
		try
		{
			file.close();
		} catch(IOException e)
		{
			e.printStackTrace();
		}
	}

	private void writeExcelSheet(Row row, Probe probe, String[] parameters, boolean amino, List<String> aminoProbes)
	{
		exWriter = new ExcelWriter(row, workbook, probe, parameters, amino, aminoProbes);
	}

	private void getExcelSheet()
	{
		try
		{
			file = new FileInputStream(path);

			workbook = new XSSFWorkbook(file);

			sheet = workbook.getSheetAt(0);
		} catch(NotOfficeXmlFileException e)
		{
			e.printStackTrace();
			JOptionPane.showMessageDialog(new JFrame(), "Die ausgewählte Datei ist keine Excel Datei!");
		} catch(Exception e)
		{
			JOptionPane.showMessageDialog(new JFrame(), "Die Datei konnte nicht geladen werden!");
			FileSaver.deleteFile(path);
			System.exit(0);
		}
	}

	public void writeSheet(XSSFWorkbook workbook)
	{
		try
		{
			FileOutputStream os = new FileOutputStream(path);
			workbook.write(os);
			os.close();
			JOptionPane.showMessageDialog(new JFrame(), "Die Datei wurde erfolgreich verarbeitet!");
		} catch(FileNotFoundException e)
		{
			JOptionPane.showMessageDialog(new JFrame(), "Die Datei wird bereits von einem anderen Prozess verwendet!");
			System.exit(0);
		} catch(Exception e)
		{
			JOptionPane.showMessageDialog(new JFrame(), "Die Datei konnte nicht beschrieben werden!");
			System.exit(0);
		}
	}

	public void removeExcelSheets(XSSFWorkbook workbook)
	{
		int noSheet = workbook.getNumberOfSheets() - 1;
		for(int i = noSheet; i > 0; i--)
		{
			workbook.removeSheetAt(i);
		}
	}

	public boolean isRowEmpty(Row row)
	{
		if(row == null)
		{
			return true;
		}
		for(int i = row.getFirstCellNum(); i < row.getLastCellNum(); ++i)
		{
			Cell cell = row.getCell(i);
			if(cell == null || cell.getCellTypeEnum() == CellType._NONE)
			{
				return true;
			}
		}
		return false;
	}
}