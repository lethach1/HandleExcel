package com.example.HandleExcel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;

import static com.example.HandleExcel.DateUtils.formateDate;

@Service
public class ExcelService {

    private Workbook getWorkbook(FileInputStream inputStream, String excelFilePath)
            throws IOException {
        Workbook workbook = null;

        if (excelFilePath.endsWith("xlsx")) {
            workbook = new XSSFWorkbook(inputStream);
        } else if (excelFilePath.endsWith("xls")) {
            workbook = new HSSFWorkbook(inputStream);
        } else {
            throw new IllegalArgumentException("The specified file is not Excel file");
        }

        return workbook;
    }

    public static ImportOrderDTO uploadSmallExcel(ExcelUploadRequest request) throws IOException {

        ImportOrderDTO dto = new ImportOrderDTO();
        SalesDataDTO salesDataDTO = new SalesDataDTO();
        List<SalesDataDTO> listDataDTO = new ArrayList<>();
        Map<String, String> mapStoreCode = new HashMap<>();

        boolean foundStoreCode = false;
        boolean hasFirstImportLine = false;

        MultipartFile file = request.getFile();

        InputStream inputStream = file.getInputStream();
        Workbook workbook = new XSSFWorkbook(inputStream);

        if (file.getOriginalFilename().endsWith("xlsx")) {
            workbook = new XSSFWorkbook(inputStream);
        } else if (file.getOriginalFilename().endsWith("xls")) {
            workbook = new HSSFWorkbook(inputStream);
        } else {
            throw new IllegalArgumentException("The specified file is not Excel file");
        }

        DataFormatter formatter = new DataFormatter();
        Sheet sheet1 = workbook.getSheetAt(0);

        for (Row row : sheet1) {
            if (row.getRowNum() == 0) continue;                 // Skip the first row which is the number column
            if (isEmptyRow(row)) continue;                      // Skip the Empty row

            if(!foundStoreCode)







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

