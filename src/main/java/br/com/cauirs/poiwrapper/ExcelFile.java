package br.com.cauirs.poiwrapper;

import static br.com.cauirs.poiwrapper.utils.Utils.generateRandomHash;
import static br.com.cauirs.poiwrapper.utils.Utils.openInputStream;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.Workbook;

import br.com.cauirs.poiwrapper.exceptions.POIException;

public class ExcelFile {

	private ExcelVersion version = ExcelVersion.XLS;
	private InputStream streamFile = null;
	private Workbook workbook = null;
	
	/**
	 * Create empty document
	 */
	public ExcelFile(ExcelVersion version) {
		this.version = version;
		workbook = version.createEmpty();
	}
	
	public ExcelFile(InputStream inputStream, ExcelVersion version) throws POIException {
		this.version = version;
		this.streamFile = inputStream;
		open();
	}

	public Workbook open() throws POIException 
	{
		if( workbook == null )
		{
			try {
				workbook = version.createFromStream(streamFile);
			} catch( FileNotFoundException ex ) {
				throw new POIException("O arquivo não foi encontrado.");
			} catch( IOException ex ) {
				throw new POIException("Houve um erro ao manipular o arquivo.");
			}
			
		} else {
			throw new POIException("O documento já está aberto.");
		}
			
		return workbook;
	}
	
	public void save(String destination) throws POIException 
	{
		try {
			OutputStream arquivoSaida = new FileOutputStream(destination);
			workbook.write(arquivoSaida);
		} catch( IOException ex ) {
			throw new POIException("Houve um erro ao manipular o arquivo.");
		}
		
	}
	
	public void createSheet() throws POIException {
		if(workbook != null) {
			workbook.createSheet();
		} else {
			throw new POIException("Não foi possível criar a planilha, pois o documento ainda não está aberto.");
		}
	}
	
	public void renameSheet(int posicaoPlanilha, String nome) throws POIException
	{
		if( workbook != null) {
			workbook.setSheetName(posicaoPlanilha, nome);
		} else {
			throw new POIException("Não foi possível renomear a planilha, pois o documento ainda não está aberto.");
		}
	}
	
	public ExcelSheet getSheet(String sheetName) throws POIException
	{
		if( workbook != null) {
			return new ExcelSheet( workbook.getSheet(sheetName) );
		} else {
			throw new POIException("Não foi possível abrir a planilha, pois o documento ainda não está aberto.");
		}	
	}
	
	public ExcelSheet getSheet(int sheetPosition) throws POIException
	{
		if( workbook != null) {
			return new ExcelSheet( workbook.getSheetAt(sheetPosition) );
		} else {
			throw new POIException("Não foi possível abrir a planilha, pois o documento ainda não está aberto.");
		}	
	}
	
	public byte[] toByteArray(String tempDir) throws POIException
	{
		StringBuilder localizacaoTemporaria = new StringBuilder(tempDir)
			.append("/")
			.append(generateRandomHash(15))
			.append(".xls");
		
		try
		{	
			OutputStream localParaSalvar = new FileOutputStream(localizacaoTemporaria.toString());
			workbook.write(localParaSalvar);
			
			InputStream localParaAbrir = openInputStream(localizacaoTemporaria.toString());
			return IOUtils.toByteArray(localParaAbrir);
		} catch( IOException ex ) 
		{
			ex.printStackTrace();
			throw new POIException("Houve um erro ao manipular o arquivo.");
		}
	}
	
}