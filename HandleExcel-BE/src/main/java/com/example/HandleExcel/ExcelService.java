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
        Map<String, String> mapStoreCode = new HashMap<>();

        boolean foundStoreCode = false;
        boolean hasFirstImportLine = false;
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
        DataFormatter formatter = new DataFormatter();
        Sheet sheet1 = workbook.getSheetAt(0);                // Get the first sheet

        int sheetCount = workbook.getNumberOfSheets();
        if (sheetCount == 0) {
            // No sheet exists in the workbook
            errors.setErrorFlg(true);
            // Error message appended
            errors.setErrorMessage(ActionMessages.GLOBAL_MESSAGE);
            return dto;
        }

        // Check exist sheet name Budget
        boolean isExistSheetNameBudget = false;
        // Check exist sheet name Budget
        boolean isExistSheetNameSales = false;
        // Check exist sheet name Budget
        boolean isExistSheetNameTC = false;

        for(int i = 0; i < sheetCount; i++){
            Sheet sheet = workbook.getSheetAt(i);
            String sheetName = sheet.getSheetName();

            // Remove spaces at the beginning and end of sheet name
            sheetName = CommonLogic.trimSheetName(sheetName);

            // To normalize sheet names by converting full-size characters to half-size
            sheetName = CommonLogic.convertToHalfSize(sheetName);
        }

        if (CommonLogic.PLAN_RESULT_FLAG.equals(request.getPlanFlag())){

        };

        for (Row row : sheet1) {
            if (row.getRowNum() == 0) continue;                 // Skip the first row which is the number column
            if (isEmptyRow(row)) continue;                      // Skip the Empty row

            if(!foundStoreCode){
                foundStoreCode = isStoreCodeRow(row, evaluator);
            }

            for (Cell cell : row) {
                CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
                System.out.print(cellRef.formatAsString());
                System.out.print(" - ");
                // get the text that appears in the cell by getting the cell value and applying any data formats (Date, 0.00, 1.23e9, $1.23, etc)
                String text = formatter.formatCellValue(cell);
                System.out.println(text);
                // Alternatively, get the value and format it yourself
                switch (cell.getCellType()) {
                    case CellType.STRING:
                        System.out.println(cell.getRichStringCellValue().getString());
                        break;
                    case CellType.NUMERIC:
                        if (DateUtil.isCellDateFormatted(cell)) {
                            System.out.println(cell.getDateCellValue());
                        } else {
                            System.out.println(cell.getNumericCellValue());
                        }
                        break;
                    case CellType.BOOLEAN:
                        System.out.println(cell.getBooleanCellValue());
                        break;
                    case CellType.FORMULA:
                        System.out.println(cell.getCellFormula());
                        break;
                    case CellType.BLANK:
                        System.out.println();
                        break;
                    default:
                        System.out.println();
                }
            }
        }
        workbook.close();
        inputStream.close();
        return dto;
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
    private static boolean isStoreCodeRow(Row row, FormulaEvaluator evaluator){
        Cell cell = row.getCell(Integer.parseInt(CommonLogic.START_COLUMN)-1);
        if(cell == null) return false;
        String value = getCellValue(cell, evaluator, false);
        return !value.isEmpty();    //check if the cell is not empty -> this is storeCode
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

