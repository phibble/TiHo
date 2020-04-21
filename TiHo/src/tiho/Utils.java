package tiho;

public class Utils
{

	public static String getFileExtention(String name)
	{
		int pointIndex = name.lastIndexOf('.');
		
		if(pointIndex < 0)
		{
			return null;
		}
		
		if(pointIndex == name.length())
		{
			return null;
		}
		
		return name.substring(pointIndex + 1, name.length()); 
	}
}