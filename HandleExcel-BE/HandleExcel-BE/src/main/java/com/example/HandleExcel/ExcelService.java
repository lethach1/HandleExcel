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
                // Thiết lập danh sách SalesDataDTO vào đối tượng trả về
                dto.setListSalesDataDTO(listDataDTO);
                break;
            };

            // Actual result flag -> read sheet Sales and TC
            if(CommonLogic.PLAN_RESULT_FLAG.equals(request.getPlanFlag())){
                if (CommonLogic.validateSheetName(sheetName, CommonLogic.SUFFIX_SALES)) {
                    isExistSheetNameSales = true;
                    listDataDTO = getListFromExcel(sheet, evaluator, errors);
                    // Thiết lập danh sách SalesDataDTO vào đối tượng trả về
                    dto.setListSalesDataDTO(listDataDTO);
                    break;
                }
                if (CommonLogic.validateSheetName(sheetName, CommonLogic.SUFFIX_TC)){
                    isExistSheetNameSales = true;
                    listDataDTO = getListFromExcel(sheet, evaluator, errors);
                    // Thiết lập danh sách SalesDataDTO vào đối tượng trả về
                    dto.setListSalesDataDTO(listDataDTO);
                    break;
                }
            }
        }

        workbook.close();
        inputStream.close();

        return dto;
    }

    // Variable to track consecutive blank rows across method calls
    private int numRowBlank = 0;

    public List<SalesDataDTO> getListFromExcel(Sheet sheet, FormulaEvaluator evaluator, Errors errors) {

        List<SalesDataDTO> listDataDTO = new ArrayList<>();
        boolean hasFirstImportLine = false;
        boolean foundStoreCode = false;

        Map<String, String> mapStoreCode = new HashMap<>();

        int endRow = Integer.parseInt(CommonLogic.END_ROW);

        // Reset blank row counter at the start of processing a new sheet
        numRowBlank = 0;

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

            // Process data row and check if we need to return empty ArrayList
            if (hasFirstImportLine) {
                if (!processDataRow(row, evaluator, sheet, errors, mapStoreCode, listDataDTO, hasFirstImportLine)) {
                    // If processDataRow returns false due to 2 consecutive blank rows, return empty ArrayList
                    return new ArrayList<>();
                }
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

        // If processBusinessDate returns null, it means we have 2 consecutive blank rows
        // Return false to signal that we should return an empty ArrayList
        if (date == null) {
            return false;
        }

        // Skip empty date
        if (date.isEmpty()) {
            return true;
        }

        // Process cells in the row and add to listDataDTO
        return processCellsInRow(row, evaluator, sheet, mapStoreCode, date, errors, listDataDTO);
    }

    private String processBusinessDate(Row row, FormulaEvaluator evaluator,
                                       Errors errors, Sheet sheet) {
        boolean isDate = true;  //cause this is process-BusinessDate method
        Cell cellDate = row.getCell(Integer.parseInt(CommonLogic.DATE_COLUMN) - 1);
        String date = getCellValueSafely(cellDate, evaluator, isDate);

        if (date.isEmpty()) {
            numRowBlank++;
            if (numRowBlank == 2){
                // Return null to signal that we should return an empty ArrayList
                return null;
            }
            return "";
        }

        if (!CommonLogic.isValidDate(date, CommonLogic.FORMAT_DATE_YYYYMMDD)) {
            // Handle invalid date
            // Note: The original code had a condition that seemed incorrect
            // It was checking if the date IS valid but then handling it as invalid
            return null;
        }

        // Reset the blank row counter when we find a valid date
        numRowBlank = 0;
        return date;
    }

    private boolean processCellsInRow(Row row, FormulaEvaluator evaluator,
                                      Sheet sheet, Map<String, String> mapStoreCode, String date,
                                      Errors errors, List<SalesDataDTO> listDataDTO) {

        // Tạo một đối tượng SalesDataDTO mới
        SalesDataDTO salesDataDTO = new SalesDataDTO();

        // Thiết lập ngày kinh doanh
        salesDataDTO.setBusinessDate(date);

        // Duyệt qua các ô trong hàng
        for (int cellIndex = Integer.parseInt(CommonLogic.START_COLUMN) - 1;
             cellIndex < Integer.parseInt(CommonLogic.END_COLUMN); cellIndex++) {
            Cell cell = row.getCell(cellIndex);
            String cellValue = getCellValueSafely(cell, evaluator, false);

            if (CheckUtil.isEmpty(cellValue)) continue;  // Bỏ qua ô trống

            // Lấy mã cửa hàng từ mapStoreCode
            String storeCode = mapStoreCode.get(String.valueOf(cellIndex + 1));
            if (storeCode != null && !storeCode.isEmpty()) {
                salesDataDTO.setStoreCode(storeCode);
            }

            // Chuyển đổi giá trị chuỗi thành BigDecimal nếu là số
            if (CommonLogic.isNumeric(cellValue)) {
                BigDecimal numericValue = new BigDecimal(cellValue);

                // Dựa vào vị trí cột để xác định loại dữ liệu
                // Đây là ví dụ, bạn cần điều chỉnh theo cấu trúc Excel của bạn
                switch (cellIndex) {
                    case 3: // Giả sử cột 3 là salesAmount
                        salesDataDTO.setSalesAmount(numericValue);
                        break;
                    case 4: // Giả sử cột 4 là pax
                        salesDataDTO.setPax(numericValue);
                        break;
                    case 5: // Giả sử cột 5 là salesBudget
                        salesDataDTO.setSalesBudget(numericValue);
                        break;
                    case 6: // Giả sử cột 6 là paxBudget
                        salesDataDTO.setPaxBudget(numericValue);
                        break;
                    // Thêm các trường hợp khác tùy theo cấu trúc Excel
                    default:
                        // Xử lý các cột khác nếu cần
                        break;
                }
            } else {
                // Xử lý các giá trị không phải số (ví dụ: weather)
                // Đây là ví dụ, bạn cần điều chỉnh theo cấu trúc Excel của bạn
                if (cellIndex == 7) { // Giả sử cột 7 là weather
                    salesDataDTO.setWeather(cellValue);
                }
            }
        }

        // Thêm đối tượng SalesDataDTO vào danh sách
        listDataDTO.add(salesDataDTO);

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
