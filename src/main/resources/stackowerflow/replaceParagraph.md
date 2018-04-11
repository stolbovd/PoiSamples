
There is the `replaceParagraph` implementation that replace `${key}` with `value` (the `fieldsForReport` parameter) and saves format by merging `runs` contents `${key}`.

<!-- language-all: lang-java -->

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

[Implementation replaceParagraph][2]

[Unit test][2]

  [1]: https://github.com/stolbovd/PoiSamples/blob/master/src/main/java/ru/inkontext/poi/SDPOIDocxView.java
  [2]: https://github.com/stolbovd/PoiSamples/blob/master/src/test/java/ru/inkontext/poi/SDPOIDocxViewTest.java