package br.com.cauirs.poiwrapper;

public class ExcelCell 
{

	private int row;	
	private String column = "";
	
	/**
	 * Construir a célula do excel. Ex: A1, A3, B6, ..
	 */
	public ExcelCell(String cellname) 
	{	
		String lineAsString = ""; //assim posso concactenar vários digitos. Ex: "1" + "2" + "3"
		
		for(char character : cellname.toCharArray())
		{
			if( Character.isLetter(character) )
				this.column += character;
			else
				lineAsString += character;
		}
		
		//como a linha é um numero, crio ela em inteiro
		this.row = new Integer(lineAsString);
	}
	
	/**
	 * 
	 * @return excel is 1-based. POI is 0-based. Return in POI`s format.
	 */
	public int getRow() 
	{
		return row - 1;
	}
	
	public int getColumn() 
	{
		char[] characters = column.toCharArray();
		
		if(characters == null || characters.length == 0)
			return -1;
		
		int columnIndex = 0;
		
		for(int i = 0; i < characters.length - 1; i++)
		{
			columnIndex += (26 * getLetterNumber(characters[i]));
		}
		
		columnIndex += getLetterNumber(characters[characters.length-1]);
		
		return columnIndex -1;
	}
	
	private int getLetterNumber(char letter)
	{
		if(Character.isLetter(letter))
			return Character.toUpperCase(letter) - 64;
		else
			return -1;
	}
	
	public String getColumnAsString()
	{
		return column;
	}
	
}