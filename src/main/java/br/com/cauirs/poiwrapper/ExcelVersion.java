package br.com.cauirs.poiwrapper;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public enum ExcelVersion {

	XLS {
		@Override
		public Workbook createEmpty() {
			return new HSSFWorkbook();
		}

		@Override
		public Workbook createFromStream(InputStream stream) throws IOException {
			return new HSSFWorkbook(new POIFSFileSystem(stream));
		}
	}, XLSX {
		@Override
		public Workbook createEmpty() {
			return new XSSFWorkbook();	
		}

		@Override
		public Workbook createFromStream(InputStream stream) throws IOException {
			return new XSSFWorkbook(stream);
		}
	};
	
	public abstract Workbook createEmpty();
	
	public abstract Workbook createFromStream(InputStream stream) throws IOException;
	
}