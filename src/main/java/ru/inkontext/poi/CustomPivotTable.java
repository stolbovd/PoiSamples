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

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.*;

import java.util.List;
import java.util.Optional;

/**
 * Created by stolbovd on 12/11/2016.
 */
public class CustomPivotTable {
	private XSSFPivotTable pivotTable;
	private CTPivotTableDefinition pivotTableDefinition;
	private CTPivotCacheDefinition pivotCacheDefinition;
	private Long lastRowIndex;


	public CustomPivotTable(XSSFSheet sheet, String source, String place) {
		pivotTable = sheet.createPivotTable(
				new AreaReference(source, SpreadsheetVersion.EXCEL2007),
				new CellReference(place));
		pivotTableDefinition = pivotTable
				.getCTPivotTableDefinition();
		pivotCacheDefinition = pivotTable
				.getPivotCacheDefinition()
				.getCTPivotCacheDefinition();
	}

	public CustomPivotTable setFormatDataField(long fieldIndex, long numFmtId) {
		Optional.ofNullable(pivotTableDefinition
				.getDataFields())
				.map(CTDataFields::getDataFieldList)
				.map(List::stream)
				.ifPresent(stream -> stream
						.filter(dataField -> dataField.getFld() == fieldIndex)
						.findFirst()
						.ifPresent(dataField -> dataField.setNumFmtId(numFmtId)));
		return this;
	}

	public CustomPivotTable setFormatPivotField(long fieldIndex, Integer numFmtId) {
		Optional.ofNullable(pivotTableDefinition
				.getPivotFields())
				.map(pivotFields -> pivotFields
						.getPivotFieldArray((int) fieldIndex))
				.ifPresent(pivotField -> pivotField
						.setNumFmtId(numFmtId));
		return this;
	}

	public CustomPivotTable addRowLabel(int fieldIndex) {
		pivotTable.addRowLabel(fieldIndex);
		return this;
	}

	public CustomPivotTable addColumnLabel(DataConsolidateFunction function, int fieldIndex) {
		pivotTable.addColumnLabel(function, fieldIndex);
		return this;
	}

	public CustomPivotTable addColLabel(int fieldIndex) {

		AreaReference pivotArea = new AreaReference(pivotCacheDefinition
				.getCacheSource()
				.getWorksheetSource()
				.getRef(), SpreadsheetVersion.EXCEL2007);
		int lastRowIndex = pivotArea.getLastCell().getRow() - pivotArea.getFirstCell().getRow();
		int lastColIndex = pivotArea.getLastCell().getCol() - pivotArea.getFirstCell().getCol();
		if (fieldIndex > lastColIndex)
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
		pivotFields.setPivotFieldArray(fieldIndex, pivotField);

		CTColFields colFields;
		if (pivotTable.getCTPivotTableDefinition().getColFields() != null) {
			colFields = pivotTable.getCTPivotTableDefinition().getColFields();
		} else {
			colFields = pivotTable.getCTPivotTableDefinition().addNewColFields();
		}

		colFields.addNewField().setX(fieldIndex);
		colFields.setCount(colFields.sizeOfFieldArray());
		return this;
	}

	// unsafe implement of exclude Subtotal
	public CustomPivotTable excludeSubTotal(int fieldIndex) {
		CTPivotField pivotField = pivotTableDefinition
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

		CTSharedItems sharedItems = pivotCacheDefinition
				.getCacheFields()
				.getCacheFieldArray(fieldIndex)
				.getSharedItems();
		sharedItems.addNewS().setV(" ");
		sharedItems.addNewS().setV("  ");

		pivotField.setDefaultSubtotal(false);
		return this;
	}

	public CustomPivotTable safeExcludeSubTotal(XSSFPivotTable pivotTable, int fieldIndex) {
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
		return this;
	}

}
