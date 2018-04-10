package ru.inkontext.poi;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import static org.junit.Assert.*;
import static ru.inkontext.poi.SDPOIDocxView.hasText;

public class SDPOIDocxViewTest {

	@Before
	public void setUp() throws Exception {
	}

	@After
	public void tearDown() throws Exception {
	}

	@Test
	public void main() {
		createDocx("result", new POITemplateView());

		assertTrue(hasText(document, "Привет"));
		assertTrue(hasText(document, "Ячейка таблицы"));
		assertTrue(hasText(document, "мир id = 5"));
		assertFalse(hasText(document, "world"));
	}
}