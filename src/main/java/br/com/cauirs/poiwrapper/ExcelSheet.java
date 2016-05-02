package br.com.cauirs.poiwrapper;

import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ExcelSheet {

	private Sheet hssfSheet;
	
	public ExcelSheet( Sheet sheet )
	{
		this.hssfSheet = sheet;
	}
	
	public Cell getCelula(String cellname)
	{
		ExcelCell cell = new ExcelCell(cellname);
		return getCell(cell);
	}
	
	public Cell getCell(ExcelCell cell) 
	{	
		Row excelRow = hssfSheet.getRow(cell.getRow());
		if( excelRow == null )
			excelRow = hssfSheet.createRow(cell.getRow());
		
		Cell excelCell = excelRow.getCell(cell.getColumn());
		
		if( excelCell == null )
			excelCell = excelRow.createCell(cell.getColumn());
			
		return excelCell;
	}
	
	public void setCellValue(ExcelCell celula, String content) 
	{
		Cell celulaDoExcel = getCell(celula);
		celulaDoExcel.setCellValue(content);
	}
	
	public void setValorCelula(ExcelCell celula, int content) 
	{
		Cell celulaDoExcel = getCell(celula);
		celulaDoExcel.setCellValue(content);
	}
	
	public void setValorCelula(ExcelCell celula, Date content) 
	{
		if( isValidContent(content) )
		{
			Cell celulaDoExcel = getCell(celula);
			celulaDoExcel.setCellValue(content);
		}
	}

	public void setValorCelula(ExcelCell celula, Double content) 
	{
		if( isValidContent(content) )
		{
			Cell celulaDoExcel = getCell(celula);
			celulaDoExcel.setCellValue(content);
		}
	}
	
	public void setValorCelula(ExcelCell celula, Boolean option) 
	{
		Cell celulaDoExcel = getCell(celula);
		celulaDoExcel.setCellValue(option);
	}
	
	private boolean isValidContent(Object conteudo)
	{
		return conteudo != null;
	}

	public void removeFormula(ExcelCell celula) 
	{
		Cell celulaDoExcel = getCell(celula);
		
		try
		{
			if( celulaDoExcel.getCellFormula() != null ) 
				celulaDoExcel.setCellFormula(null);
		}
		catch(Exception ex)
		{
			//Do nothing.
		} 
	}
	
	@Override
	public String toString() {
		if(hssfSheet != null)
			return hssfSheet.getSheetName();
		
		return "";
	}
}