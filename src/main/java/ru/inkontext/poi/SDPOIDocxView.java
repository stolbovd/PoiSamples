package ru.inkontext.poi;

import org.apache.poi.POIXMLException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

public abstract class SDPOIDocxView {

	protected XWPFDocument document;

	public static void main(String[] args) throws IOException, InvalidFormatException {
		InputStream fis = SDPOIDocxView.class
				.getClassLoader()
				.getResourceAsStream("templates/docx/template.docx");

		XWPFDocument xdoc = new XWPFDocument(OPCPackage.open(fis));

		SDPOIDocxView.replace(xdoc, "${Hello}", "Привет");
		SDPOIDocxView.replace(xdoc, "${world}", "мир");
		SDPOIDocxView.replace(xdoc, "${Table cell}", "Ячейка таблицы");

		xdoc.write(new FileOutputStream("result.docx"));
	}

	private void renderDocument(XWPFDocument document) throws IOException {
		document.write(out);
		document.close();
	}

	protected abstract void build() throws Exception;

	public static List<String> getTagValues(final String str) {
		final List<String> tagValues = new ArrayList<String>();
		Pattern pattern = Pattern.compile("\\$\\{(.*?)\\}");
		Matcher matcher = pattern.matcher(str);
		while (matcher.find()) {
			tagValues.add(matcher.group());
		}

		return tagValues;
	}

	public static List<XWPFRun> findInParagraph(XWPFParagraph paragraph, String findText) {
		return paragraph.getRuns().stream()
				.filter(run -> Optional.ofNullable(run
						.getText(0)).map(text -> text
						.contains(findText))
						.orElse(false))
				.collect(Collectors.toList());
	}

	public static void replaceRuns(XWPFParagraph paragraph, String findText, String replaceText) {
		paragraph.getRuns().forEach(run -> Optional.ofNullable(run
				.getText(0)).ifPresent(text -> {
			if (text.contains(findText)) {
				text = text.replace(findText, replaceText);
				run.setText(text, 0);
			}
		}));
	}

	public static void iterateParagraphs(XWPFDocument doc, Consumer<XWPFParagraph> consumer) {
		for (XWPFParagraph p : doc.getParagraphs())
			consumer.accept(p);
		for (XWPFTable tbl : doc.getTables())
			for (XWPFTableRow row : tbl.getRows())
				for (XWPFTableCell cell : row.getTableCells())
					for (XWPFParagraph p : cell.getParagraphs())
						consumer.accept(p);
	}

	public static Boolean hasParagraphs(XWPFDocument doc, Function<XWPFParagraph, Boolean> function) {
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

	public static List<XWPFRun> find(XWPFDocument doc, String findText) {
		List<XWPFRun> runs = new ArrayList<>();
		iterateParagraphs(doc, p -> runs.addAll(findInParagraph(p, findText)));
		return runs;
	}

	public static Boolean hasText(XWPFDocument doc, String findText) {
		return hasParagraphs(doc, p -> p.getText().contains(findText));
	}

	public static void replace(XWPFDocument doc, String findText, String replaceText) {
		iterateParagraphs(doc, p -> replaceRuns(p, findText, replaceText));
	}

	protected List<XWPFRun> find(String findText) {
		return find(this.document, findText);
	}

	protected void replace(String findText, String replaceText) {
		replace(this.document, findText, replaceText);
	}

	protected void replace(Map<String, String> fieldsForReport) {
		iterateParagraphs(this.document, p -> replaceParagraph(p, fieldsForReport));
	}

	private static void replaceParagraph(XWPFParagraph paragraph, Map<String, String> fieldsForReport) throws POIXMLException {
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
