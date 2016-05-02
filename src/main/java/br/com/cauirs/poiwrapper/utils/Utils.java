package br.com.cauirs.poiwrapper.utils;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.Random;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import com.google.common.base.Strings;

import br.com.cauirs.poiwrapper.exceptions.POIException;

public class Utils {

	public static String generateRandomHash(int length) {
		char[] chars =
			{ '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' };

			char[] hash = new char[length];

			int chartLenght = chars.length;
			Random rdm = new Random();

			for (int x = 0; x < length; x++)
				hash[x] = chars[rdm.nextInt(chartLenght)];

			return new String(hash);
	}
	
	public static InputStream openInputStream(String localizacaoArquivo) throws POIException
	{
		try
		{
			if(localizacaoArquivo.contains("http"))
			{
				URL enderecoDoArquivo = new URL(localizacaoArquivo);
				return enderecoDoArquivo.openStream();
			} 
			else 
			{
				return new FileInputStream(localizacaoArquivo);
			}
		}
		catch(IOException ex)
		{
			throw new POIException("Erro ao abrir o arquivo: " + localizacaoArquivo);
		}
	}
	
	public static String onlyNumbers(String word)
	{
		if (Strings.isNullOrEmpty(word)) return "";

		String patternString = "[^0-9]*";
		String replaceStr = "";
		Pattern pattern = Pattern.compile(patternString);
		Matcher matcher = pattern.matcher(word);

		return matcher.replaceAll(replaceStr);
	}
	
}