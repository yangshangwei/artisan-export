package com.artisan.service;


import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public interface CircuitServiceInterface {
	
	public XSSFWorkbook exportFormAndCircuitInfo(String flowId) throws Exception;
}
