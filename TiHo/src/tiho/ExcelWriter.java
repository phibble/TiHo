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

		if(workbook.getSheet(valueID) == null)
		{
			newSheet = workbook.createSheet(valueID);
			countNewRows = 0;
		} else
		{
			newSheet = workbook.getSheet(valueID);
			countNewRows = newSheet.getLastRowNum() + 1;
		}

		Row newRow = null;

		if(countNewRows != 0)
		{
			newSheet.createRow(countNewRows++); // blank line but not on the first line
		}

		newRow = newSheet.createRow(countNewRows++);

		newRow.createCell(0).setCellValue(probe.getNumber());
		newRow.createCell(1).setCellValue(probe.getName());

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

			if(workbook.getSheet(valueID) == null)
			{
				newSheet = workbook.createSheet(valueID);
				countNewRows = 0;
			} else
			{
				newSheet = workbook.getSheet(valueID);
				countNewRows = newSheet.getLastRowNum() + 1;
			}

			Row newRow = null;

			if(countNewRows != 0)
			{
				newSheet.createRow(countNewRows++); // blank line but not on the first line
			}

			newRow = newSheet.createRow(countNewRows++);

			newRow.createCell(0).setCellValue(probe.getNumber());
			newRow.createCell(1).setCellValue(probe.getName());

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
		if(row.getCell(index).getCellTypeEnum() == CellType.STRING)
		{
			newRow.createCell(index).setCellValue(row.getCell(index).getStringCellValue());
		} else if(row.getCell(index).getCellTypeEnum() == CellType.NUMERIC)
		{
			newRow.createCell(index).setCellValue(row.getCell(index).getNumericCellValue());
		}
	}
}