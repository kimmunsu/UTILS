package com.darkship.util;

public enum POIFiletypes {
	XLS("xls"),
	XLSX("xlsx");
	
	private String str;
	
	private POIFiletypes(String arg){
		this.str = arg;
	}
	public String get(){
		return str;
	}
}
