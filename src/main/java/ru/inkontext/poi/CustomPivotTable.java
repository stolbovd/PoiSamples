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
import org.openxmlformats.schemas.spreadsheetml.x2006.main.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Optional;

public class CustomPivotTable {

	public static void main(String[] args) throws IOException, InvalidFormatException {
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet();

		//Create some data to build the pivot table on
		setCellData(sheet);

		XSSFPivotTable pivotTable = sheet.createPivotTable(
				new AreaReference("A1:C6", SpreadsheetVersion.EXCEL2007),
				new CellReference("E3"));

		// set first column as 1-th level of rows
		pivotTable.addRowLabel(0);
		// excude subtotal
//		excludeSubTotal(pivotTable, 0);
		safeExcludeSubTotal(pivotTable, 0);
		// set second column of source as 2-th level of rows
		pivotTable.addRowLabel(1);
		// Sum up the second column
		pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 2);
		setFormatDataField(pivotTable, 2, 4); //# ##0.00

		FileOutputStream fileOut = new FileOutputStream("custom-pivottable.xlsx");
		wb.write(fileOut);
		fileOut.close();
		wb.close();
	}

	private static void setFormatDataField(XSSFPivotTable pivotTable, long fieldIndex, long numFmtId) {
		Optional.ofNullable(pivotTable.getCTPivotTableDefinition().getDataFields())
				.map(CTDataFields::getDataFieldList)
				.map(List::stream)
				.ifPresent(stream -> stream
						.filter(dataField -> dataField.getFld() == fieldIndex)
						.findFirst()
						.ifPresent(dataField -> dataField.setNumFmtId(numFmtId)));
	}

	// unsafe implement of exclude Subtotal
	private static void excludeSubTotal(XSSFPivotTable pivotTable, int fieldIndex) {
		CTPivotField pivotField = pivotTable
				.getCTPivotTableDefinition()
				.getPivotFields()
				.getPivotFieldArray(fieldIndex);

		CTItems items = pivotField.getItems();
		for (int i = 0; i < 2; i++) {
			items.getItemArray(i).unsetT();
			items.getItemArray(i).setX((long) i);
		}
		for (int i = items.sizeOfItemArray() - 1; i > 1; i--)
			items.removeItem(i);
		items.setCount(2);

		CTSharedItems sharedItems = pivotTable.getPivotCacheDefinition()
				.getCTPivotCacheDefinition()
				.getCacheFields()
				.getCacheFieldArray(fieldIndex)
				.getSharedItems();
		sharedItems.addNewS().setV(" ");
		sharedItems.addNewS().setV("  ");

		pivotField.setDefaultSubtotal(false);
	}

	private static void setCellData(XSSFSheet sheet) {

		String[] cities = {"Rome", "Paris", "Rome", "Paris", "Athens"};
		String[] names = {"Jane", "Tarzan", "Terk", "Kate", "Dmitry"};
		Integer[] balances = {10, 5, 10, 20, 15};

		Row row = sheet.createRow(0);
		row.createCell(0).setCellValue("City");
		row.createCell(1).setCellValue("Name");
		row.createCell(2).setCellValue("Balance");

		for (int i = 0; i < cities.length; i++) {
			row = sheet.createRow(i + 1);
			row.createCell(0).setCellValue(cities[i]);
			row.createCell(1).setCellValue(names[i]);
			row.createCell(2).setCellValue(balances[i]);
		}
	}

	private static void addColLabel(XSSFPivotTable pivotTable, int columnIndex) {
		AreaReference pivotArea = new AreaReference(pivotTable.getPivotCacheDefinition().getCTPivotCacheDefinition()
				.getCacheSource().getWorksheetSource().getRef(), SpreadsheetVersion.EXCEL2007);
		int lastRowIndex = pivotArea.getLastCell().getRow() - pivotArea.getFirstCell().getRow();
		int lastColIndex = pivotArea.getLastCell().getCol() - pivotArea.getFirstCell().getCol();
		if (columnIndex > lastColIndex)
			throw new IndexOutOfBoundsException();

		CTPivotFields pivotFields = pivotTable.getCTPivotTableDefinition().getPivotFields();
		CTPivotField pivotField = CTPivotField.Factory.newInstance();
		CTItems items = pivotField.addNewItems();

		pivotField.setAxis(STAxis.AXIS_COL);
		pivotField.setShowAll(false);
		for (int i = 0; i <= lastRowIndex; i++) {
			items.addNewItem().setT(STItemType.DEFAULT);
		}
		items.setCount(items.sizeOfItemArray());
		pivotFields.setPivotFieldArray(columnIndex, pivotField);

		CTColFields rowFields;
		if (pivotTable.getCTPivotTableDefinition().getColFields() != null) {
			rowFields = pivotTable.getCTPivotTableDefinition().getColFields();
		} else {
			rowFields = pivotTable.getCTPivotTableDefinition().addNewColFields();
		}

		rowFields.addNewField().setX(columnIndex);
		rowFields.setCount(rowFields.sizeOfFieldArray());
	}

	private static void safeExcludeSubTotal(XSSFPivotTable pivotTable, int fieldIndex) {
		Optional.ofNullable(pivotTable.getCTPivotTableDefinition().getPivotFields())
				.map(pivotFields -> pivotFields.getPivotFieldArray(fieldIndex))
				.ifPresent(pivotField ->
						Optional.ofNullable(pivotField.getItems())
								.ifPresent(items -> {
									for (int i = 0; i < 2; i++) {
										items.getItemArray(i).unsetT();
										items.getItemArray(i).setX((long) i);
									}
									for (int i = items.sizeOfItemArray() - 1; i > 1; i--)
										items.removeItem(i);
									items.setCount(2);

									Optional.ofNullable(pivotTable.getPivotCacheDefinition()
											.getCTPivotCacheDefinition().getCacheFields())
											.map(CTCacheFields::getCacheFieldArray)
											.ifPresent(ctCacheFields ->
													Optional.ofNullable(ctCacheFields[fieldIndex])
															.map(CTCacheField::getSharedItems)
															.ifPresent(ctSharedItems -> {
																ctSharedItems.addNewS().setV(" ");
																ctSharedItems.addNewS().setV("  ");
															}));

									pivotField.setDefaultSubtotal(false);
								}));
	}
}
