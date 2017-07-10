/**
 * 
 */
package com.arun.util;

import java.io.File;
import java.io.IOException;
import java.nio.file.Paths;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.xssf.extractor.XSSFBEventBasedExcelExtractor;
import org.apache.poi.xssf.extractor.XSSFEventBasedExcelExtractor;
import org.apache.xmlbeans.XmlException;

/**
 * @author Adwiti
 *
 */
public class Util {

	/**
	 * 
	 * @param fileName
	 * @param access
	 * @return
	 * 
	 * 		The method opens an xlsb file and reading
	 */
	public static XSSFEventBasedExcelExtractor openAnXLSBFile(String fileName, PackageAccess access) {
		File file = null;
		OPCPackage pkg = null;
		XSSFBEventBasedExcelExtractor evt = null;

		file = Paths.get(fileName).toFile();
		try {
			pkg = OPCPackage.open(file, access);
			ZipSecureFile.setMaxTextSize(110485766);
			evt = new XSSFBEventBasedExcelExtractor(pkg);
		} catch (XmlException | OpenXML4JException | IOException e) {
			e.printStackTrace();
		}

		return evt;
	}

	public static void readAnXLSBFile(XSSFEventBasedExcelExtractor extractor) {
		System.out.println(extractor.getText());
	}

}
