This is **NOSAFE** (without checks of `null`, `IndexOutOfBounds`, ...) implements of **exclude Subtotal** from **Pivot table Field**.


Thanks for answer in [Apache POI XSSFPivotTable setDefaultSubtotal][2] by [Axel Richter][3] resolved this task.

<!-- language-all: lang-java -->

    // NOSAFE implement of exclude Subtotal
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

        // should be get real data from source of pivotTable by fieldIndex
        sharedItems.addNewS().setV(" ");
        sharedItems.addNewS().setV("  ");
    
        pivotField.setDefaultSubtotal(false);
    }

Below is example calling `excludeSubTotal`, based on [official POI sample CreatePivotTable][1]:

    public static void main(String[] args) throws IOException, InvalidFormatException {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet();

        //Create some data to build the pivot table on
        fillCellData(sheet);

        XSSFPivotTable pivotTable = sheet.createPivotTable(
            new AreaReference("A1:C6", SpreadsheetVersion.EXCEL2007),
            new CellReference("E3"));

        // set first column as 1-th level of rows
        pivotTable.addRowLabel(0);
        // excude subtotal
        excludeSubTotal(pivotTable, 0);
        // set second column of source as 2-th level of rows
        pivotTable.addRowLabel(1);
        // Sum up the second column
        pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 2);

        FileOutputStream fileOut = new FileOutputStream("ooxml-pivottable.xlsx");
        wb.write(fileOut);
        fileOut.close();
        wb.close();
    }

fill Cell Data:

    private static void fillCellData(XSSFSheet sheet) {

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

and **SAFE** implements for *turn off subtotals* with **Java8** `Optional`

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

[1]: https://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/xssf/usermodel/examples/CreatePivotTable.java
[2]: http://stackoverflow.com/questions/37305976/apache-poi-xssfpivottable-setdefaultsubtotal?answertab=active#tab-top
[3]: http://stackoverflow.com/users/3915431/axel-richter