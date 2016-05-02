package br.com.cauirs.poiwrapper;

import java.util.Date;
import org.apache.poi.ss.usermodel.Cell;
import com.google.common.base.Strings;
import static br.com.cauirs.poiwrapper.utils.Utils.*;

public class ExcelExtrator
{

	public ExcelExtrator(ExcelSheet sheet)
	{
		this.sheet = sheet;
	}

	private final ExcelSheet sheet;

	public String getString(String nomeCelula)
	{
		try
		{
			return sheet.getCelula(nomeCelula).getStringCellValue();
		}
		catch (NullPointerException ex)
		{
			return "";
		}
	}

	public String getFormula(String nomeCelula)
	{
		try
		{
			String formula = sheet.getCelula(nomeCelula).getCellFormula();

			return formula;
		}
		catch (NullPointerException ex)
		{
			return getCellType(nomeCelula) == 0 ? String
					.valueOf(getInteger(nomeCelula)) : getString(nomeCelula);
		}
	}

	public Integer getInteger(String nomeCelula)
	{
		try
		{
			return (int) sheet.getCelula(nomeCelula).getNumericCellValue();
		}
		catch (NullPointerException ex)
		{
			return 0;
		}
	}

	public Date getDate(String nomeCelula)
	{
		try
		{
			return sheet.getCelula(nomeCelula).getDateCellValue();
		}
		catch (NullPointerException ex)
		{
			return null;
		}
	}

	public Double getNumeric(String nomeCelula)
	{
		try
		{
			return sheet.getCelula(nomeCelula).getNumericCellValue();
		}
		catch (NullPointerException ex)
		{
			return new Double(0.0);
		}
	}

	public Boolean getBoolean(String nomeCelula)
	{
		return sheet.getCelula(nomeCelula).getBooleanCellValue();
	}

	public int getCellType(String nomeCelula)
	{
		try
		{
			return sheet.getCelula(nomeCelula).getCellType();
		}
		catch (NullPointerException ex)
		{
			return -1;
		}
	}

	public String getStringOrNumberAsString(String nomeCelula)
	{
		try
		{
			sheet.getCelula(nomeCelula).setCellType(Cell.CELL_TYPE_STRING);
			return getString(nomeCelula).trim();

		}
		catch (Exception e)
		{
			return getCellType(nomeCelula) == 0 ? String
					.valueOf(getInteger(nomeCelula)) : getString(nomeCelula);
		}

	}

	public Double getMoney(String nomeCelula)
	{
		if (getCellType(nomeCelula) == 0 || getCellType(nomeCelula) == 2)
		{
			return getNumeric(nomeCelula);
		}

		String valorCelula = getString(nomeCelula);
		String numero = onlyNumbers(valorCelula);

		if (!Strings.isNullOrEmpty(numero))
		{
			if (valorCelula.contains(",") || valorCelula.contains("."))
			{
				return Double.parseDouble(numero) / 100;
			}
			else
			{
				return Double.parseDouble(numero);
			}
		}

		return 0.0;
	}
	
}