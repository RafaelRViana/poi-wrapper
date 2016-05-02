package br.com.cauirs.poiwrapper.exceptions;

import java.io.IOException;

public class POIException extends IOException {

	private static final long serialVersionUID = 297467485813956621L;

	public POIException(String message) {
		super(message);
	}
	
}