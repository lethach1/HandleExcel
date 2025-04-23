package com.example.HandleExcel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.*;

import static com.example.HandleExcel.DateUtils.formateDate;

@Service
public class ExcelService {

    public ImportOrderStatus uploadSmallExcel(ExcelUploadRequest request) throws IOException {

        ImportOrderStatus dto = new ImportOrderStatus();
        SalesDataDTO salesDataDTO = new SalesDataDTO();
        List<SalesDataDTO> listDataDTO = new ArrayList<>();

        Errors errors = new Errors();

        MultipartFile file = request.getFile();
        InputStream inputStream = file.getInputStream();
        Workbook workbook = null;

        if (file.getOriginalFilename().endsWith("xlsx")) {
            workbook = new XSSFWorkbook(inputStream);
        } else if (file.getOriginalFilename().endsWith("xls")) {
            workbook = new HSSFWorkbook(inputStream);
        } else {
            throw new IllegalArgumentException("The specified file is not Excel file");
        }

        // Create FormulaEvaluator formulas in Workbook
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

        // Check exist sheet name Budget
        boolean isExistSheetNameBudget = false;
        // Check exist sheet name Sales
        boolean isExistSheetNameSales = false;
        // Check exist sheet name TC
        boolean isExistSheetNameTC = false;

        int sheetCount = workbook.getNumberOfSheets();
        if (sheetCount == 0) {
            // No sheet exists in the workbook
            errors.setErrorFlg(true);
            // Error message appended
            errors.setErrorMessage(ActionMessages.GLOBAL_MESSAGE);
            return dto;
        }

        // Read sheet list in file excel
        for(int i = 0; i < sheetCount; i++){
            Sheet sheet = workbook.getSheetAt(i);
            String sheetName = sheet.getSheetName();

            // Remove spaces at the beginning and end of sheet name
            sheetName = CommonLogic.trimSheetName(sheetName);

            // To normalize sheet names by converting full-size characters to half-size
            sheetName = CommonLogic.convertToHalfSize(sheetName);

            //Plan result flag -> only read sheet Sales
            if (CommonLogic.PLAN_RESULT_FLAG.equals(request.getPlanFlag())
                    && CommonLogic.validateSheetName(sheetName, CommonLogic.SUFFIX_BUDGET)) {
                isExistSheetNameBudget = true;
                listDataDTO = getListFromExcel(sheet, evaluator, errors);
                return dto;
            };

            // Actual result flag -> read sheet Sales and TC
            if(CommonLogic.PLAN_RESULT_FLAG.equals(request.getPlanFlag())){
                if (CommonLogic.validateSheetName(sheetName, CommonLogic.SUFFIX_SALES)) {
                    isExistSheetNameSales = true;
                }
                if (CommonLogic.validateSheetName(sheetName, CommonLogic.SUFFIX_TC)){
                    isExistSheetNameSales = true;

                }
            }
        }

        workbook.close();
        inputStream.close();

        return dto;
    }

    public List<SalesDataDTO> getListFromExcel(Sheet sheet, FormulaEvaluator evaluator, Errors errors) {

        List<SalesDataDTO> listDataDTO = new ArrayList<>();
        boolean hasFirstImportLine = false;
        boolean foundStoreCode = false;

        Map<String, String> mapStoreCode = new HashMap<>();

        int endRow = Integer.parseInt(CommonLogic.END_ROW);

        //process each row
        for(int rowIndex = 0; rowIndex < endRow; rowIndex++){
            if (rowIndex == 0) continue;                        // Skip the first row which is the number column

            Row row = sheet.getRow(rowIndex);
            if (isEmptyRow(row)) continue;                      // Skip the Empty row

            // Process store code row
            if (!foundStoreCode){
                if (!processStoreCodeRow(row, evaluator, mapStoreCode, errors, sheet)) {
                    continue;
                }
                foundStoreCode = true;
                continue;
            }

            if(foundStoreCode && !hasFirstImportLine){
                if(isDataLine(row, evaluator)){
                    hasFirstImportLine = true;
                    endRow= rowIndex + CommonLogic.MAX_NUMBER_OF_DAYS_PER_YEAR;
                } else {
                    continue;
                }
            }
        }
        return listDataDTO;
    }



    private boolean processDataRow(Row row, FormulaEvaluator evaluator, Sheet sheet, Errors errors,
                                   Map<String, String> mapStoreCode, List<SalesDataDTO> listDataDTO, boolean hasFirstImportLine) {

        // Handle first data line detection
        if (!hasFirstImportLine) {
            if (isDataLine(row, evaluator)) {
                hasFirstImportLine = true;
                int endRow= row.getRowNum() + CommonLogic.MAX_NUMBER_OF_DAYS_PER_YEAR;
            } else {
                return true;
            }
        }

        // Process business date
        String date = processBusinessDate(row, evaluator, errors, sheet);
        if (date == null) return false;
    }

    private String processBusinessDate(Row row, FormulaEvaluator evaluator,
                                       Errors errors, Sheet sheet) {

        Cell cellDate = row.getCell(Integer.parseInt(CommonLogic.DATE_COLUMN) - 1);
        String date = getCellValueSafely(cellDate, evaluator, true);

        if (date.isEmpty()) {
            return handleEmptyDate(state, errorFlg, sheet, row);
        }

        if (CommonLogic.isValidDate(date, CommonLogic.FORMAT_DATE_YYYYMMDD)) {
            handleInvalidDate(errorFlg, sheet, date, row);
            return null;
        }

        state.numRowBlank = 0;
        return date;
    }


    private boolean processStoreCodeRow(Row row, FormulaEvaluator evaluator,
                                        Map<String, String> mapStoreCode, Errors errors, Sheet sheet) {
        if (!isStoreCodeRow(row, evaluator)) {
            return false;
        }

        // if found Store Code row, read all Store Code data
        // Iterate through each cell in the row
        for (int cellIndex = Integer.parseInt(CommonLogic.START_COLUMN) - 1;
             cellIndex < Integer.parseInt(CommonLogic.END_COLUMN); cellIndex++) {
            //Get value from cell
            String cellValue = getCellValueSafely(row, cellIndex, evaluator );

            if (cellValue.isEmpty()) continue;                  // Skip empty cell

            String errPrefix = " [row " + (row.getRowNum() + 1) + "] " + " [col " + (cellIndex + 1) + "] ";

            mapStoreCode.put(String.valueOf(cellIndex + 1), cellValue);
        }

        return true;
    }




    private boolean isEmptyRow(Row row){
        if(row == null) return true;
        for (Cell cell : row){
            if(cell != null && cell.getCellType() != CellType.BLANK){
                return false;
            }
        }
        return true;
    }

    //check if the cell is storeCode
    private boolean isStoreCodeRow(Row row, FormulaEvaluator evaluator){
        Cell cell = row.getCell(Integer.parseInt(CommonLogic.START_COLUMN)-1);
        if(cell == null) return false;
        String value = getCellValue(cell, evaluator, false);
        return !value.isEmpty();    //check if the cell is not empty -> this is storeCode
    }



    /**
     * Detect if is row contains import date by detect type of value of column "Date".
     * @return true, if value column "Date" has date format, otherwise return false.
     */
    private boolean isDataLine(Row row, FormulaEvaluator evaluator){
        // Check if the business date column (Column A) contains a valid date
        Cell dateCell = row.getCell(Integer.parseInt(CommonLogic.DATE_COLUMN) - 1);
        if (dateCell == null) return false;
        return DateUtil.isCellDateFormatted(dateCell);
    }


    private String getCellValueSafely(Row row, int cellIndex, FormulaEvaluator evaluator) {
        Cell cell = row.getCell(cellIndex);
        return !(cell == null) ? getCellValue(cell, evaluator, false) : "";
    }

    private String getCellValue(Cell cell, FormulaEvaluator evaluator, boolean isDate){
        //Handle Formula
        CellValue cellValue = evaluator.evaluate(cell);
        if(cellValue == null) return "";

        switch (cell.getCellType()) {
            case CellType.NUMERIC:
                if (DateUtil.isCellDateFormatted(cell) && isDate) {
                    return formateDate(cell.getDateCellValue(), CommonLogic.FORMAT_DATE_YYYYMMDD);
                } else {
                    return new BigDecimal(cellValue.getNumberValue()).toPlainString();
                }
            case CellType.STRING:
               return cell.getStringCellValue().trim();
            case CellType.BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case CellType.BLANK:
                return "";
            default:
                return "";
        }

    }
}

