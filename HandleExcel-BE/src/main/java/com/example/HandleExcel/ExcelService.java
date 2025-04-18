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

    public static ImportOrderStatus uploadSmallExcel(ExcelUploadRequest request) throws IOException {

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
                listDataDTO = getListFromExcel();
                return dto;
            };

            // Actual result flag -> read sheet Sales and TC
            if(CommonLogic.PLAN_RESULT_FLAG.equals(request.getPlanFlag()){
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

    public static List<SalesDataDTO> getListFromExcel(Sheet sheet, FormulaEvaluator evaluator, Errors errors) {

        boolean hasFirstImportLine = false;
        Map<String, String> mapStoreCode = new HashMap<>();

        int endRow = Integer.parseInt(CommonLogic.END_ROW);

        for(int rowIndex = 0; rowIndex < endRow; rowIndex++){
            if (rowIndex == 0) continue;                 // Skip the first row which is the number column
            Row row = sheet.getRow(rowIndex);
            if (isEmptyRow(row)) continue;                      // Skip the Empty row

            // if found Store Code row, read all Store Code data
            boolean foundStoreCode = isStoreCodeRow(row, evaluator);
            if(foundStoreCode){
                // Iterate through each cell in the row
                for (int cellIndex = Integer.parseInt(CommonLogic.START_COLUMN) - 1;
                     cellIndex < Integer.parseInt(CommonLogic.END_COLUMN); cellIndex++) {

                    int rowNum = rowIndex + 1;
                    int colNum = cellIndex + 1;
                    String errPrefix = " [row " + rowNum + "] " + " [col " + colNum + "] ";

                    Cell cell = row.getCell(cellIndex);

                    //Get value from cell
                    String cellValue = "";
                    if (cell != null) {
                        cellValue = getCellValue(cell, evaluator, false);
                    }

                    if (cellValue.isEmpty()) {
                        // Skip empty cell
                        continue;
                    }

                    mapStoreCode.put(String.valueOf(cellIndex + 1), cellValue);
                }
            }

            if(foundStoreCode && !hasFirstImportLine){
                if(isDataLine(row, evaluator)){

                }

            }


        }

    }




    private static boolean isEmptyRow(Row row){
        if(row == null) return true;
        for (Cell cell : row){
            if(cell != null && cell.getCellType() != CellType.BLANK){
                return false;
            }
        }
        return true;
    }

    //check if the cell is storeCode
    private static boolean isStoreCodeRow(Row row, FormulaEvaluator evaluator){
        Cell cell = row.getCell(Integer.parseInt(CommonLogic.START_COLUMN)-1);
        if(cell == null) return false;
        String value = getCellValue(cell, evaluator, false);
        return !value.isEmpty();    //check if the cell is not empty -> this is storeCode
    }

    private boolean isDataLine(Row row, FormulaEvaluator evaluator){
        Cell cell = row.getCell(Integer.parseInt(CommonLogic.STD_DATE_COLUMN) - 1);
        if (CheckUtil.isNull(cell)) return false;
        return HSSFDateUtil.isCellDateFormatted(cell);
    }

    private static String getCellValue(Cell cell, FormulaEvaluator evaluator, boolean isDate){
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

