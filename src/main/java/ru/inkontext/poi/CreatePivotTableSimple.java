/* ====================================================================
 Licensed to the Apache Software Foundation (ASF) under one or more
 contributor license agreements.  See the NOTICE file distributed with
 this work for additional information regarding copyright ownership.
 The ASF licenses this file to You under the Apache License, Version 2.0
 (the "License"); you may not use this file except in compliance with
 the License.  You may obtain a copy of the License at

 http://www.apache.org/licenses/LICENSE-2.0

 Unless required by applicable law or agreed to in writing, software
 distributed under the License is distributed on an "AS IS" BASIS,
 WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 See the License for the specific language governing permissions and
 limitations under the License.
 ==================================================================== */
package ru.inkontext.poi;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataFields;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Optional;

public class CreatePivotTableSimple {

	public static void main(String[] args) throws IOException, InvalidFormatException {

		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet();

		//Create some data to build the pivot table on
		setCellData(sheet);

		XSSFPivotTable pivotTable = sheet.createPivotTable(
				new AreaReference("A1:C6", SpreadsheetVersion.EXCEL2007),
				new CellReference("E3"));

		pivotTable.addRowLabel(1); // set second column as 1-th level of rows
		setFormatPivotField(pivotTable, 1, 9); //set format numFmtId=9 0%
		pivotTable.addRowLabel(0); // set first column as 2-th level of rows
		pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 2); // Sum up the second column
		setFormatDataField(pivotTable, 2, 3); //numFmtId=3 # ##0

		FileOutputStream fileOut = new FileOutputStream("stackoverflow-pivottable.xlsx");
		wb.write(fileOut);
		fileOut.close();
		wb.close();
	}

	private static void setCellData(XSSFSheet sheet) {

		String[] names = {"Jane", "Tarzan", "Terk", "Kate", "Dmitry"};
		Double[] percents = {0.25, 0.5, 0.75, 0.25, 0.5};
		Integer[] balances = {107634, 554234, 10234, 22350, 15234};

		Row row = sheet.createRow(0);
		row.createCell(0).setCellValue("Name");
		row.createCell(1).setCellValue("Percents");
		row.createCell(2).setCellValue("Balance");

		for (int i = 0; i < names.length; i++) {
			row = sheet.createRow(i + 1);
			row.createCell(0).setCellValue(names[i]);
			row.createCell(1).setCellValue(percents[i]);
			row.createCell(2).setCellValue(balances[i]);
		}
	}

	private static void setFormatDataField(XSSFPivotTable pivotTable,
											   long fieldIndex,
											   long numFmtId) {
		Optional.ofNullable(pivotTable
				.getCTPivotTableDefinition()
				.getDataFields())
				.map(CTDataFields::getDataFieldList)
				.map(List::stream)
				.ifPresent(stream -> stream
						.filter(dataField -> dataField.getFld() == fieldIndex)
						.findFirst()
						.ifPresent(dataField -> dataField.setNumFmtId(numFmtId)));
	}

	private static void setFormatPivotField(XSSFPivotTable pivotTable,
												long fieldIndex,
												Integer numFmtId) {
		Optional.ofNullable(pivotTable
				.getCTPivotTableDefinition()
				.getPivotFields())
				.map(pivotFields -> pivotFields
						.getPivotFieldArray((int) fieldIndex))
				.ifPresent(pivotField -> pivotField
						.setNumFmtId(numFmtId));
	}

}
