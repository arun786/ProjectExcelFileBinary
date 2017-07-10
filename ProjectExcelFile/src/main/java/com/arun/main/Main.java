/**
 * 
 */
package com.arun.main;

import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.extractor.XSSFEventBasedExcelExtractor;

import com.arun.util.Util;

/**
 * @author Adwiti
 *
 */
public class Main {
	public static void main(String[] args) {
		XSSFEventBasedExcelExtractor ext = Util.openAnXLSBFile("C:/Users/Adwiti/Desktop/Book1.xlsb", PackageAccess.READ_WRITE);
		Util.readAnXLSBFile(ext);
	}

}
