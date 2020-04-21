package tiho;

import java.io.File;

import javax.swing.filechooser.FileFilter;

public class XLSXFilter extends FileFilter
{
	
	public boolean accept(File f)
	{
		if(f.isDirectory())
		{
			return true;
		}
		
		String name = f.getName();
		String extension = Utils.getFileExtention(name);
		
		if(extension == null)
		{
			return false;
		}
		if(extension.trim().toLowerCase().equals("xlsx"))
		{
			return true;
		}
		
		return false;
	}
	
	public String getDescription()
	{
		return "Excel Files (*.xlsx)";
	}
}