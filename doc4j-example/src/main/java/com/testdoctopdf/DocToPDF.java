package com.testdoctopdf;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.docx4j.convert.out.pdf.viaXSLFO.PdfSettings;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFont;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

public class DocToPDF {

	public static void main(String[] args) {
		
		try {
			
			InputStream is = new FileInputStream(
					new File(
							"/Users/badalb/projects/docx4j/doc4j-example/src/main/resources/doc/MERGEFIELD-OUT.docx"));
			WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage
					.load(is);
			// If your header and body information got over lapped then use the
			// below code
			List<SectionWrapper> sections = wordMLPackage.getDocumentModel().getSections();
			for (int i = 0; i < sections.size(); i++) {

				wordMLPackage.getDocumentModel().getSections().get(i)
						.getPageDimensions();
			}
		
			Mapper fontMapper = new IdentityPlusMapper();
			PhysicalFont font = PhysicalFonts.getPhysicalFonts().get(
					"Comic Sans MS");
			fontMapper.getFontMappings().put("Algerian", font);
			wordMLPackage.setFontMapper(fontMapper);

			// 2) Prepare Pdf settings
			PdfSettings pdfSettings = new PdfSettings();

			// 3) Convert WordprocessingMLPackage to Pdf

			org.docx4j.convert.out.pdf.PdfConversion conversion = new org.docx4j.convert.out.pdf.viaXSLFO.Conversion(
					wordMLPackage);

			OutputStream out = new FileOutputStream(
					new File(
							"/Users/badalb/projects/docx4j/doc4j-example/src/main/resources/doc/MERGEFIELD.pdf"));
			conversion.output(out, pdfSettings);
		} catch (Throwable e) {

			e.printStackTrace();
		} 
	}

}
