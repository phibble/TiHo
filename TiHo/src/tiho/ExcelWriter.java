package tiho;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
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
	private boolean amino;
	private List<String> aminoProbes;

	public ExcelWriter(Row row, XSSFWorkbook workbook, Probe probe, String[] parameters, boolean amino,
			List<String> aminoProbes)
	{
		this.row = row;
		this.workbook = workbook;
		this.probe = probe;
		this.amino = amino;
		this.aminoProbes = aminoProbes;

		writeExcel(parameters);
	}

	private void writeExcel(String[] parameters)
	{
		Cell cell = row.getCell(0);
		valueID = cell.getStringCellValue().trim();

		List<String> paramList = null;
		if(parameters != null)
		{
			paramList = convertArrayToList(parameters);
		}

		if(amino)
		{
			writeAminoAcids();
		} else if(paramList == null || paramList.contains(valueID)
				|| ((valueID.toLowerCase().contains("ts")) ? paramList.contains("TS") : false))
		{
			int countNewRows = 0;

			XSSFSheet newSheet = null;

			if(valueID.toLowerCase().contains("ts"))
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

				copyNumericStringValue(row, i, newRow, row.getCell(i).getCellStyle());
			}
		}
	}

	private void writeAminoAcids()
	{
		XSSFSheet aminoSheet = null;
		int aminoRowCounter = 0;

		if(workbook.getSheet("Aminosäuren") == null)
		{
			aminoSheet = workbook.createSheet("Aminosäuren");
		} else
		{
			aminoSheet = workbook.getSheet("Aminosäuren");
			aminoRowCounter = aminoSheet.getLastRowNum();
		}

		Row aminoRow = null;
		Cell cell = null;

		if(!aminoProbes.contains(probe.getName()))
		{
			aminoRow = aminoSheet.createRow(aminoRowCounter++);

			cell = aminoRow.createCell(0);
			cell.setCellValue(probe.getNumber());
			cell = aminoRow.createCell(1);
			cell.setCellValue(probe.getName());
		}

		aminoRow = aminoSheet.createRow(aminoRowCounter++);

		for(int i = 0; i < row.getLastCellNum(); i++)
		{
			if(row.getCell(i) == null)
			{
				continue;
			}

			copyNumericStringValue(row, i, aminoRow, row.getCell(i).getCellStyle());
		}

		aminoRow = aminoSheet.createRow(aminoRowCounter++);
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

	public void copyNumericStringValue(Row row, int index, Row newRow, CellStyle cellStyle)
	{
		Cell cell = row.getCell(index);
		
		if(cell.getCellTypeEnum() == CellType.STRING)
		{
			cell = newRow.createCell(index);
			cell.setCellValue(row.getCell(index).getStringCellValue());
			cell.setCellStyle(cellStyle);
		} else if(cell.getCellTypeEnum() == CellType.NUMERIC)
		{
			cell = newRow.createCell(index);
			cell.setCellValue(row.getCell(index).getNumericCellValue());
			cell.setCellStyle(cellStyle);
		}
	}
}