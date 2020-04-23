package tiho;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
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

	public static List<String> readPrevFile()
	{
		createPath();
		createPrevFile();

		List<String> prev = new ArrayList<String>();

		try
		{
			File file = new File(path);
			Scanner reader = new Scanner(file);

			while(reader.hasNextLine())
			{
				prev.add(reader.nextLine());
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
			List<String> prevText = readPrevFile();
			FileWriter writer = new FileWriter(path);

			for(int i = 0; i < prevText.size(); i++)
			{
				writer.write(prevText.get(i));
				writer.write("\n");
			}

			writer.write(text);
			writer.close();
		} catch(IOException e)
		{
			e.printStackTrace();
		}
	}

	public List<String> changeLastFile()
	{
		List<String> prevText = readPrevFile();
		List<String> result = new ArrayList<String>();

		if(prevText.indexOf(text) != prevText.size() - 1)
		{
			for(String str : prevText)
			{
				if(!str.equals(text))
				{
					result.add(str);
				}
			}
			result.add(text);
		}

		return result;
	}

	public static void deleteFile(String text)
	{
		try
		{
			List<String> prevText = readPrevFile();
			FileWriter writer = new FileWriter(path);
			
			System.out.println("*" + text + "*");
			
			for(int i = 0; i < prevText.size(); i++)
			{
				if(!prevText.get(i).trim().equals(text))
				{
					System.out.println("*" + prevText.get(i) + "*");
					writer.write(prevText.get(i));
					writer.write("\n");
				}
			}
			
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