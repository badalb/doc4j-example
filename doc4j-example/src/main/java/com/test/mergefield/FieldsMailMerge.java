package com.test.mergefield;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.docx4j.model.fields.merge.DataFieldName;
import org.docx4j.model.fields.merge.MailMerger.OutputField;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

public class FieldsMailMerge {

	public static void main(String[] args) throws Exception {
		

		boolean mergedOutput = false;
		
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(
				new java.io.File("/Users/badalb/projects/docx4j/doc4j-example/src/main/resources/doc/MERGEFIELD.docx"));
		
		List<Map<DataFieldName, String>> data = new ArrayList<Map<DataFieldName, String>>();

		// Instance 1
		Map<DataFieldName, String> map = new HashMap<DataFieldName, String>();
		map.put( new DataFieldName("KundenNAme"), "Daffy duck");
		map.put( new DataFieldName("Kundenname"), "Plutext");
		map.put(new DataFieldName("Kundenstrasse"), "Bourke Street");
		map.put(new DataFieldName("yourdate"), "15/4/2013");  
		map.put(new DataFieldName("yournumber"), "2456800");
		data.add(map);
				
			
		if (mergedOutput) {
		
			// How to treat the MERGEFIELD, in the output?
			org.docx4j.model.fields.merge.MailMerger.setMERGEFIELDInOutput(OutputField.KEEP_MERGEFIELD);
			
			WordprocessingMLPackage output = org.docx4j.model.fields.merge.MailMerger.getConsolidatedResultCrude(wordMLPackage, data, true);
			
			output.save(new java.io.File("/Users/badalb/projects/docx4j/doc4j-example/src/main/resources/doc/MERGEFIELD-OUT.docx") );
			
		} else {

			org.docx4j.model.fields.merge.MailMerger.setMERGEFIELDInOutput(OutputField.KEEP_MERGEFIELD);			
			for (Map<DataFieldName, String> thismap : data) {
				org.docx4j.model.fields.merge.MailMerger.performMerge(wordMLPackage, thismap, true);
				wordMLPackage.save(new java.io.File("/Users/badalb/projects/docx4j/doc4j-example/src/main/resources/doc/MERGEFIELD-OUT.docx") );
			}			
		}
		
	}

}
