package com.example.HandleExcel.Service;

import com.example.HandleExcel.Utils.CheckUtil;
import com.example.HandleExcel.Utils.CommonLogic;
import com.example.HandleExcel.DTO.ExcelUploadRequest;
import com.example.HandleExcel.DTO.ImportOrderStatus;
import com.example.HandleExcel.DTO.SalesDataDTO;
import com.example.HandleExcel.Errors;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;


import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.*;

import static com.example.HandleExcel.Utils.DateUtils.formateDate;

@Service
public class ExcelService {

    @Setter @Getter
    private static class ProcessingState {
        int endRow;
        boolean foundStoreCode = false;
        boolean hasFirstImportLine = false;
        int numRowBlank = 0;
    }

    /* Helper class to track sheet existence */
    @Getter @Setter
    private static class SheetValidator {
        private boolean budgetExists = false;
        private boolean salesExists = false;
        private boolean tcExists = false;
    }

    public ImportOrderStatus uploadSmallExcel(ExcelUploadRequest request) throws IOException {
        ImportOrderStatus dto = new ImportOrderStatus();
        Errors errors = new Errors();

        MultipartFile file = request.getFile();
        InputStream inputStream = file.getInputStream();
        String fileName = file.getOriginalFilename();

        Workbook workbook = createWorkbook(inputStream,fileName);
        if (workbook == null) {
            dto.setListSalesDataDTO(new ArrayList<>());
            dto.setErrors(errors);
            return dto;
        }

        // Create FormulaEvaluator formulas in Workbook
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        SheetValidator sheetValidator = new SheetValidator();
        ProcessingState processingState = new ProcessingState();

        // No sheet exists in the workbook
        if (!isSheetExist(workbook, errors)) {
            dto.setListSalesDataDTO(new ArrayList<>());
            dto.setErrors(errors);
            return dto;
        }

        // Process sheets based on plan flag
        if (CommonLogic.PLAN_RESULT_FLAG.equals(request.getPlanFlag())) {
            processPlanResult(workbook, evaluator, dto, errors, sheetValidator);
            if (!sheetValidator.budgetExists) {
                dto.setListSalesDataDTO(new ArrayList<>());
                dto.setErrors(errors);
                return dto;
            }
        } else {
            processActualResult(workbook, evaluator, dto, errors, sheetValidator);
            if (!sheetValidator.salesExists || !sheetValidator.tcExists) {
                dto.setListSalesDataDTO(new ArrayList<>());
                dto.setErrors(errors);
                return dto;
            }
        }

        inputStream.close();
        workbook.close();
        return dto;
    }


    private Workbook createWorkbook(InputStream inputStream, String fileName) throws IOException {
        if (fileName.endsWith("xlsx")) {
            return new XSSFWorkbook(inputStream);
        } else if (fileName.endsWith("xls")) {
            return new HSSFWorkbook(inputStream);
        } else {
            throw new IllegalArgumentException("The specified file is not Excel file");
        }
    }

    private boolean isSheetExist(Workbook workbook, Errors errors) {
        if (workbook.getNumberOfSheets() == 0) {
            errors.setErrorFlg(false);
            errors.addErrorMessage("No sheet exists in the workbook");
            return false;
        }
        return true;
    }

    private void processPlanResult(Workbook workbook, FormulaEvaluator evaluator,
                                   ImportOrderStatus dto, Errors errors, SheetValidator validator) {
        //get budget sheet file in excel
        Sheet budgetSheet = findSheet(workbook, CommonLogic.SUFFIX_BUDGET);
        if (budgetSheet != null) {
            validator.setBudgetExists(true);
            List<SalesDataDTO> listDataDTO = getListFromExcel(budgetSheet, evaluator, errors);
            dto.setErrors(errors);
            dto.setListSalesDataDTO(listDataDTO);
        }
    }

    private void processActualResult(Workbook workbook, FormulaEvaluator evaluator,
                                     ImportOrderStatus dto, Errors errors, SheetValidator validator) {
        //get sales sheet file in excel
        Sheet salesSheet = findSheet(workbook, CommonLogic.SUFFIX_SALES);
        if (salesSheet != null) {
            validator.setSalesExists(true);
            List<SalesDataDTO> listDataDTO = getListFromExcel(salesSheet, evaluator, errors);
            dto.setErrors(errors);
            dto.setListSalesDataDTO(listDataDTO);
        }

        //get TC sheet file in excel
        Sheet tcSheet = findSheet(workbook, CommonLogic.SUFFIX_TC);
        if (tcSheet != null) {
            validator.setTcExists(true);
            List<SalesDataDTO> listDataDTO = getListFromExcel(tcSheet, evaluator, errors);
            dto.setErrors(errors);
            dto.setListSalesDataDTO(listDataDTO);
        }
    }

    // Read sheet list in file excel
    private Sheet findSheet(Workbook workbook, String suffix) {
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            String sheetName = normalizeSheetName(sheet.getSheetName());

            if (CommonLogic.validateSheetName(sheetName, suffix)) {
                return sheet;
            }
        }
        return null;
    }

    // To normalize sheet names by converting full-size characters to half-size
    private String normalizeSheetName(String sheetName) {
        // Remove spaces at the beginning and end of sheet name
        sheetName = CommonLogic.trimSheetName(sheetName);
        return CommonLogic.convertToHalfSize(sheetName);
    }

    public List<SalesDataDTO> getListFromExcel(Sheet sheet, FormulaEvaluator evaluator, Errors errors) {

        List<SalesDataDTO> listDataDTO = new ArrayList<>();
        ProcessingState processingState = new ProcessingState();

        Map<String, String> mapStoreCode = new HashMap<>();

        int endRow = Integer.parseInt(CommonLogic.END_ROW);

        //process each row
        for(int rowIndex = 0; rowIndex < endRow; rowIndex++){
            if (rowIndex == 0) continue;                        // Skip the first row which is the number column

            Row row = sheet.getRow(rowIndex);
            if (isEmptyRow(row)) continue;                      // Skip the Empty row

            // Process store code row
            if (!processingState.foundStoreCode){
                if (!processStoreCodeRow(row, evaluator, mapStoreCode, errors, sheet)) {
                    continue;
                }
                processingState.foundStoreCode = true;
                continue;
            }

            if(processingState.foundStoreCode && !processingState.hasFirstImportLine){
                if(isDataLine(row, evaluator)){
                    processingState.hasFirstImportLine = true;
                    endRow= rowIndex + CommonLogic.MAX_NUMBER_OF_DAYS_PER_YEAR;
                } else {
                    continue;
                }
            }

            // Process data rows
            if (!processDataRow(row, evaluator, sheet, errors, mapStoreCode, listDataDTO, processingState)) {
                return new ArrayList<>();
            }
        }
        return listDataDTO;
    }


    private boolean processStoreCodeRow(Row row, FormulaEvaluator evaluator,
                                        Map<String, String> mapStoreCode, Errors errors, Sheet sheet) {
        if (!isStoreCodeRow(row, evaluator)) {
            return false;
        }

        boolean isDate = false;  //cause this is storeCode

        // if found Store Code row, read all Store Code data
        // Iterate through each cell in the row
        for (int cellIndex = Integer.parseInt(CommonLogic.START_COLUMN) - 1;
             cellIndex < Integer.parseInt(CommonLogic.END_COLUMN); cellIndex++) {
            //Get value from cell
            Cell cell = row.getCell(cellIndex);
            String cellValue = getCellValueSafely(cell, evaluator, isDate);

            if (cellValue.isEmpty()) continue;                  // Skip empty cell

            String errPrefix = " [row " + (row.getRowNum() + 1) + "] " + " [col " + (cellIndex + 1) + "] ";

            mapStoreCode.put(String.valueOf(cellIndex + 1), cellValue);
        }

        return true;
    }



    private boolean processDataRow(Row row, FormulaEvaluator evaluator, Sheet sheet, Errors errors,
                                   Map<String, String> mapStoreCode, List<SalesDataDTO> listDataDTO, ProcessingState processingState) {
        // Handle first data line detection
        if (!processingState.hasFirstImportLine) {
            if (isDataLine(row, evaluator)) {
                processingState.hasFirstImportLine = true;
                processingState.endRow = row.getRowNum() + CommonLogic.MAX_NUMBER_OF_DAYS_PER_YEAR;
            } else {
                return true;
            }
        }

        // Process business date
        String date = processBusinessDate(row, evaluator, processingState);
        if (date == null) return false;

        // Process each cell in the row
        processCellsInRow(row, evaluator, sheet, mapStoreCode, date, errors, listDataDTO);
        return true;
    }

    private String processBusinessDate(Row row, FormulaEvaluator evaluator,ProcessingState processingState) {
        boolean isDate = true;  //cause this is process-BusinessDate method


        Cell cellDate = row.getCell(Integer.parseInt(CommonLogic.DATE_COLUMN) - 1);
        String date = getCellValueSafely(cellDate, evaluator, isDate);
        if (date.isEmpty()) {
            processingState.numRowBlank++;  // Variable to track consecutive blank rows across method calls
            if (processingState.numRowBlank++ == 2){
                return null;
            }
            return "";
        }
        processingState.numRowBlank = 0;
        return date;
    }

    private void processCellsInRow(Row row, FormulaEvaluator evaluator,
                                      Sheet sheet, Map<String, String> mapStoreCode, String date,
                                      Errors errors, List<SalesDataDTO> listSalesDataDTO) {

        for (int cellIndex = Integer.parseInt(CommonLogic.START_COLUMN) - 1;
             cellIndex < Integer.parseInt(CommonLogic.END_COLUMN); cellIndex++) {
            Cell cell = row.getCell(cellIndex);
            String cellValue = getCellValueSafely(cell, evaluator, false);

            if (CheckUtil.isEmpty(cellValue)) continue;                  // Skip empty cell

            String errPrefix = " [row " + (row.getRowNum() + 1) + "] " + " [col " + (cellIndex + 1) + "] ";

            // process validate cell value each cell
            if (!isCellValueValid(cellValue, errors, errPrefix)) {
                continue;
            }
            // Add each cell to listDataDTO
            SalesDataDTO salesDataDTO = new SalesDataDTO();
            if (mapStoreCode.get(String.valueOf(cellIndex + 1)) == null ){
                errors.setErrorFlg(true);
                errors.addErrorMessage(sheet.getSheetName() + errPrefix + "Store code is empty");
                continue;
                //if store code is empty, skip this cell
            }
            salesDataDTO.setStoreCode(mapStoreCode.get(String.valueOf(cellIndex + 1)));
            salesDataDTO.setBusinessDate(date);

            //set cell value
            if (CommonLogic.validateSheetName(sheet.getSheetName(), CommonLogic.SUFFIX_BUDGET)) {
                salesDataDTO.setSalesBudget(cellValue);
            }
            if (CommonLogic.validateSheetName(sheet.getSheetName(), CommonLogic.SUFFIX_SALES)) {
                salesDataDTO.setSalesAmount(cellValue);
            }
            if (CommonLogic.validateSheetName(sheet.getSheetName(), CommonLogic.SUFFIX_TC)) {
                salesDataDTO.setPax(cellValue);
            }
            listSalesDataDTO.add(salesDataDTO);
        }
    }

    private boolean isCellValueValid(String cellValue, Errors errors, String errPrefix) {
        if (!CommonLogic.isNumeric(cellValue)) {
            errors.setErrorFlg(true);
            errors.addErrorMessage(errPrefix + "Not a number");
            return false;
        }

        if (CommonLogic.isNegativeNumeric(cellValue)) {
            errors.setErrorFlg(true);
            errors.addErrorMessage(errPrefix + "Negative number is not allowed");
            return false;
        }

        if (!CommonLogic.hasShortIntegerPart(cellValue, CommonLogic.MAX_LENGTH_NUMERIC)){
            errors.setErrorFlg(true);
            errors.addErrorMessage(errPrefix + "The number of digits is too long");
            return false;
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

    /*
     * Detect if is row contains import date by detect type of value of column "Date".
     * @return true, if value column "Date" has date format, otherwise return false.
     */
    private boolean isDataLine(Row row, FormulaEvaluator evaluator){
        // Check if the business date column (Column A) contains a valid date
        Cell dateCell = row.getCell(Integer.parseInt(CommonLogic.DATE_COLUMN) - 1);
        if (dateCell == null) return false;
        return DateUtil.isCellDateFormatted(dateCell);
    }


    private String getCellValueSafely(Cell cell, FormulaEvaluator evaluator, boolean isDate) {
        return !(cell == null) ? getCellValue(cell, evaluator, isDate) : "";
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

