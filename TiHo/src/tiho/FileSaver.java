package tiho;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Scanner;

public class FileSaver
{
	private static String path;
	
	private String text;
	
	public FileSaver(String text)
	{
		this.text = text;
		
		createPath();
		
		createPrevFile();
		readPrevFile();
	}
	
	public static void createPath()
	{
		try
		{
			String dir = new File(".").getCanonicalPath();
			path = dir.substring(0, dir.lastIndexOf("\\")) + "\\prev\\prev.txt";
		} catch(IOException e)
		{
			path = "\\prev\\prev.txt";
			e.printStackTrace();
		}
	}
	
	public static String readPrevFile()
	{
		createPath();
		createPrevFile();
		
		String prev = "";
		
		int counter = 0;
		
		try
		{
			File file = new File(path);
			Scanner reader = new Scanner(file);
			
			while(reader.hasNextLine())
			{
				prev += (counter != 0 ? "," : "") + reader.nextLine();
				counter++;
			}
			
			reader.close();
		} catch(FileNotFoundException e)
		{
			e.printStackTrace();
		}
		
		return prev;
	}
	
	public void writePrevFile()
	{
		createPath();
		createPrevFile();
		
		try
		{
			String prevText = readPrevFile();
			FileWriter writer = new FileWriter(path);
			writer.write(prevText + (prevText != "" ? "\n" : "") + text);
			writer.close();
		} catch(IOException e)
		{
			e.printStackTrace();
		}
	}
	
	private static void createPrevFile()
	{
		File dir = new File(path.substring(0, path.lastIndexOf("\\")));
		
		if(!dir.exists())
		{
			dir.mkdir();
		} else
		{
			File prevFile = new File(path);
			
			try
			{
				prevFile.createNewFile();
			} catch(IOException e)
			{
				e.printStackTrace();
			}
		}
	}
}