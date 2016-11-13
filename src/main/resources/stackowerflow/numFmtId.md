
Format of pivot table fields is setting by `CTDataField.setNumFmtId(long numFmtId)` **for values** and `CTPivotField.setNumFmtId(long numFmtId)` **for columns & rows**.

`numFmtId` is id number of format code. Available format codes are represented in Format cells list - Custom category:
[![enter image description here][1]][1]
Predefined format codes, thanks to [Ji Zhou - MSFT][2], is here:

    1 0    
    2 0.00    
    3 #,##0    
    4 #,##0.00    
    5 $#,##0_);($#,##0)    
    6 $#,##0_);[Red]($#,##0)    
    7 $#,##0.00_);($#,##0.00)    
    8 $#,##0.00_);[Red]($#,##0.00)    
    9 0%    
    10 0.00%    
    11 0.00E+00    
    12 # ?/?    
    13 # ??/??    
    14 m/d/yyyy    
    15 d-mmm-yy    
    16 d-mmm    
    17 mmm-yy    
    18 h:mm AM/PM    
    19 h:mm:ss AM/PM    
    20 h:mm    
    21 h:mm:ss    
    22 m/d/yyyy h:mm    
    37 #,##0_);(#,##0)    
    38 #,##0_);[Red](#,##0)    
    39 #,##0.00_);(#,##0.00)    
    40 #,##0.00_);[Red](#,##0.00)    
    45 mm:ss    
    46 [h]:mm:ss    
    47 mm:ss.0    
    48 ##0.0E+0    
    49 @    
    
Full list of predefined format codes in [MSDN NumberingFormat Class][3]

Here is an example of applying format pivot table fields:

<!-- language-all: lang-java -->

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

        public static void main(String[] args) throws IOException, InvalidFormatException {
    
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = wb.createSheet();
    
            //Create some data to build the pivot table on
            setCellData(sheet);
    
            XSSFPivotTable pivotTable = sheet.createPivotTable(
                    new AreaReference("A1:C6", SpreadsheetVersion.EXCEL2007),
                    new CellReference("E3"));
    
            pivotTable.addRowLabel(1); // set second column as 1-th level of rows
            setFormatPivotField(pivotTable, 1, 9); //set format of row field numFmtId=9 0%
            pivotTable.addRowLabel(0); // set first column as 2-th level of rows
            pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 2); // Sum up the second column
            setFormatDataField(pivotTable, 2, 3); //set format of value field numFmtId=3 # ##0
    
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
    }

https://github.com/stolbovd/PoiSamples

  [1]: https://i.stack.imgur.com/4EhZF.png
  [2]: https://social.msdn.microsoft.com/Forums/office/en-US/e27aaf16-b900-4654-8210-83c5774a179c/xlsx-numfmtid-predefined-id-14-doesnt-match?forum=oxmlsdk
  [3]: https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.numberingformat%28v=office.15%29.aspx?f=255&MSPPError=-2147217396