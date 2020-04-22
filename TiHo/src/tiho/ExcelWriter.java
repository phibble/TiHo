package tiho;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriter
{
	private Row row;
	private XSSFWorkbook workbook;
	private String valueID;
	private Probe probe;

	public ExcelWriter(Row row, XSSFWorkbook workbook, Probe probe)
	{
		this.row = row;
		this.workbook = workbook;
		this.probe = probe;

		writeExcel();
	}

	public ExcelWriter(Row row, XSSFWorkbook workbook, Probe probe, String[] parameters)
	{
		this.row = row;
		this.workbook = workbook;
		this.probe = probe;

		writeExcel(parameters);
	}

	private void writeExcel()
	{
		Cell cell = row.getCell(0);
		valueID = cell.getStringCellValue().trim();

		int countNewRows = 0;

		XSSFSheet newSheet = null;

		if(valueID.contains("TS"))
		{
			if(workbook.getSheet("TS") == null)
			{
				newSheet = workbook.createSheet("TS");
				countNewRows = 0;
			} else
			{
				newSheet = workbook.getSheet("TS");
				countNewRows = newSheet.getLastRowNum() + 1;
			}
		} else
		{
			if(workbook.getSheet(valueID) == null)
			{
				newSheet = workbook.createSheet(valueID);
				countNewRows = 0;
			} else
			{
				newSheet = workbook.getSheet(valueID);
				countNewRows = newSheet.getLastRowNum() + 1;
			}
		}

		Row newRow = null;

		if(!(valueID.contains("2.") && valueID.contains("TS")))
		{
			if(countNewRows != 0)
			{
				newRow = newSheet.createRow(countNewRows++); // blank line but not on the first line
			}

			newRow = newSheet.createRow(countNewRows++);

			cell = newRow.createCell(0);
			cell.setCellValue(probe.getNumber());
			cell = newRow.createCell(1);
			cell.setCellValue(probe.getName());
		}

		newRow = newSheet.createRow(countNewRows++);

		for(int i = 0; i < row.getLastCellNum(); i++)
		{
			if(row.getCell(i) == null)
			{
				continue;
			}

			copyNumericStringValue(row, i, newRow);
		}
	}

	private void writeExcel(String[] parameters)
	{
		Cell cell = row.getCell(0);
		valueID = cell.getStringCellValue().trim();

		List<String> paramList = convertArrayToList(parameters);

		if(paramList.contains(valueID))
		{
			int countNewRows = 0;

			XSSFSheet newSheet = null;

			if(valueID.contains("TS"))
			{
				if(workbook.getSheet("TS") == null)
				{
					newSheet = workbook.createSheet("TS");
					countNewRows = 0;
				} else
				{
					newSheet = workbook.getSheet("TS");
					countNewRows = newSheet.getLastRowNum() + 1;
				}
			} else
			{
				if(workbook.getSheet(valueID) == null)
				{
					newSheet = workbook.createSheet(valueID);
					countNewRows = 0;
				} else
				{
					newSheet = workbook.getSheet(valueID);
					countNewRows = newSheet.getLastRowNum() + 1;
				}
			}

			Row newRow = null;

			if(!(valueID.contains("2.") && valueID.contains("TS")))
			{
				if(countNewRows != 0)
				{
					newRow = newSheet.createRow(countNewRows++); // blank line but not on the first line
				}

				newRow = newSheet.createRow(countNewRows++);

				cell = newRow.createCell(0);
				cell.setCellValue(probe.getNumber());
				cell = newRow.createCell(1);
				cell.setCellValue(probe.getName());
			}

			newRow = newSheet.createRow(countNewRows++);

			for(int i = 0; i < row.getLastCellNum(); i++)
			{
				if(row.getCell(i) == null)
				{
					continue;
				}

				copyNumericStringValue(row, i, newRow);
			}
		}
	}

	private List<String> convertArrayToList(String[] parameters)
	{
		List<String> result = new ArrayList<String>();

		for(String str : parameters)
		{
			result.add(str);
		}

		return result;
	}

	public void copyNumericStringValue(Row row, int index, Row newRow)
	{
		Cell cell = row.getCell(index);
		if(cell.getCellTypeEnum() == CellType.STRING)
		{
			cell = newRow.createCell(index);
			cell.setCellValue(row.getCell(index).getStringCellValue());
		} else if(cell.getCellTypeEnum() == CellType.NUMERIC)
		{
			cell = newRow.createCell(index);
			cell.setCellValue(row.getCell(index).getNumericCellValue());
		}
	}
}