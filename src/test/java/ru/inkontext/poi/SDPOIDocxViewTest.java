package ru.inkontext.poi;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.Before;
import org.junit.Test;

import java.io.FileInputStream;
import java.util.HashMap;

import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;
import static ru.inkontext.poi.SDPOIDocxView.hasText;

public class SDPOIDocxViewTest {

	@Before
	public void setUp() throws Exception {
		SDPOIDocxView docxView = new SDPOIDocxView("docx/template.docx");

		docxView.replace(new HashMap<String, String>() {{
			put("Hello", "Lorem");
			put("world", "ipsum");
			put("Table cell", "Inside table");
		}});

		docxView.writeAndClose("result.docx");
	}

	@Test
	public void replaceParagraphTest() throws Exception {

		XWPFDocument document = new XWPFDocument(new FileInputStream("result.docx"));

		assertTrue(hasText(document, "Lorem"));
		assertTrue(hasText(document, "Inside table"));
		assertTrue(hasText(document, "ipsum"));
		assertFalse(hasText(document, "world"));

		document.close();
	}
}