package ru.inkontext.poi;

import org.apache.poi.POIXMLException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;
import java.util.function.Function;

class SDPOIDocxView {

	private XWPFDocument document;

	SDPOIDocxView(String templatePath) throws IOException, InvalidFormatException {
		document = openDocument(templatePath);
	}

	private XWPFDocument openDocument(String filePath) throws InvalidFormatException, IOException {
		InputStream fis = SDPOIDocxView.class
				.getClassLoader()
				.getResourceAsStream(filePath);
		return new XWPFDocument(OPCPackage.open(fis));
	}

	void writeAndClose(String fileResult) throws IOException {
		document.write(new FileOutputStream(fileResult));
		document.close();
	}

	private void iterateParagraphs(XWPFDocument doc, Consumer<XWPFParagraph> consumer) {
		for (XWPFParagraph p : doc.getParagraphs())
			consumer.accept(p);
		for (XWPFTable tbl : doc.getTables())
			for (XWPFTableRow row : tbl.getRows())
				for (XWPFTableCell cell : row.getTableCells())
					for (XWPFParagraph p : cell.getParagraphs())
						consumer.accept(p);
	}

	private static Boolean hasParagraphs(XWPFDocument doc, Function<XWPFParagraph, Boolean> function) {
		for (XWPFParagraph p : doc.getParagraphs())
			if (function.apply(p))
				return true;
		for (XWPFTable tbl : doc.getTables())
			for (XWPFTableRow row : tbl.getRows())
				for (XWPFTableCell cell : row.getTableCells())
					for (XWPFParagraph p : cell.getParagraphs())
						if (function.apply(p))
							return true;
		return false;
	}

	static Boolean hasText(XWPFDocument doc, String findText) {
		return hasParagraphs(doc, p -> p.getText().contains(findText));
	}

	void replace(Map<String, String> fieldsForReport) {
		iterateParagraphs(document, p -> replaceParagraph(p, fieldsForReport));
	}

	private void replaceParagraph(XWPFParagraph paragraph, Map<String, String> fieldsForReport) throws POIXMLException {
		String find, text, runsText;
		List<XWPFRun> runs;
		XWPFRun run, nextRun;
		for (String key : fieldsForReport.keySet()) {
			text = paragraph.getText();
			if (!text.contains("${"))
				return;
			find = "${" + key + "}";
			if (!text.contains(find))
				continue;
			runs = paragraph.getRuns();
			for (int i = 0; i < runs.size(); i++) {
				run = runs.get(i);
				runsText = run.getText(0);
				if (runsText.contains("${")) {
					while (!runsText.contains("}")) {
						nextRun = runs.get(i + 1);
						runsText = runsText + nextRun.getText(0);
						paragraph.removeRun(i + 1);
					}
					run.setText(runsText.contains(find) ?
							runsText.replace(find, fieldsForReport.get(key)) :
							runsText, 0);
				}
			}
		}
	}
}
