package com.example.exportdatatoexcel.utils;

import com.example.exportdatatoexcel.entity.Customer;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;
import org.springframework.util.ObjectUtils;
import org.springframework.util.ResourceUtils;

import java.io.*;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

import static com.example.exportdatatoexcel.utils.FileFactory.PATH_TEMPLATE;

@Slf4j
@Component
public class ExcelUtils {

    //export config
    public static ByteArrayInputStream exportCustomer(List<Customer> customers, String fileName) throws Exception {

        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();

        //get file -> not found -> create file
        File file;
        FileInputStream fileInputStream;

        try {
            file = ResourceUtils.getFile(PATH_TEMPLATE + fileName);
            fileInputStream = new FileInputStream(file);
        } catch (Exception e) {
            log.info("FILE NOT FOUND");
            file = FileFactory.createFile(fileName, xssfWorkbook);
            fileInputStream = new FileInputStream(file);
        }

        processInsertData(xssfWorkbook, customers);

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        xssfWorkbook.write(outputStream);

        //close resource
        outputStream.close();
        fileInputStream.close();

        log.info("done");
        return new ByteArrayInputStream(outputStream.toByteArray());
    }

    public static List<File> getFilesExcelStoreDataFromDatabase(List<Customer> customers, String fileName) throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook();

        File file;
        try {
            file = ResourceUtils.getFile(PATH_TEMPLATE + fileName);
        }catch (Exception e){
            log.info("File not found");
            file = FileFactory.createFile(fileName, workbook);
        }

        processInsertData(workbook, customers);

        FileOutputStream fileOutputStream = new FileOutputStream(file);
        workbook.write(fileOutputStream);
        List<File> result = new ArrayList<>();
        result.add(file);
        return result;
    }

    private static void processInsertData(XSSFWorkbook xssfWorkbook, List<Customer> customers){
        //create freeze pane in excel file
        XSSFSheet newSheet = xssfWorkbook.createSheet("sheet1");
        newSheet.createFreezePane(4, 2, 4, 2);


        //create font for title
        XSSFFont titleFont = xssfWorkbook.createFont();
        titleFont.setFontName("Arial");
        titleFont.setBold(true);
        titleFont.setFontHeightInPoints((short) 14);


        //create style for cell of title and apply font to cell
        XSSFCellStyle titleCellStyle = xssfWorkbook.createCellStyle();
        titleCellStyle.setAlignment(HorizontalAlignment.CENTER);
        titleCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        titleCellStyle.setFillForegroundColor(IndexedColors.SKY_BLUE.index);
        titleCellStyle.setBorderBottom(BorderStyle.MEDIUM);
        titleCellStyle.setBorderLeft(BorderStyle.MEDIUM);
        titleCellStyle.setBorderRight(BorderStyle.MEDIUM);
        titleCellStyle.setBorderTop(BorderStyle.MEDIUM);
        titleCellStyle.setFont(titleFont);
        titleCellStyle.setWrapText(true);


        //create font for data
        XSSFFont dataFont = xssfWorkbook.createFont();
        dataFont.setFontName("Arial");
        dataFont.setBold(false);
        dataFont.setFontHeightInPoints((short) 10);

        //create style for cell data and apply font data to cell
        XSSFCellStyle dataCellStyle = xssfWorkbook.createCellStyle();
        dataCellStyle.setAlignment(HorizontalAlignment.CENTER);
        dataCellStyle.setBorderBottom(BorderStyle.THIN);
        dataCellStyle.setBorderLeft(BorderStyle.THIN);
        dataCellStyle.setBorderRight(BorderStyle.THIN);
        dataCellStyle.setBorderTop(BorderStyle.THIN);
        dataCellStyle.setFont(dataFont);
        dataCellStyle.setWrapText(true);

        //insert fieldName as title to excel
        insertFieldNameAsTitleToWorkbook(ExportConfig.customerExport.getCellExportConfigList(), newSheet, titleCellStyle);


        //insert data of fieldName to excel
        insertDataToWorkbook(xssfWorkbook, ExportConfig.customerExport, customers, dataCellStyle);
        //return
    }

    private static <T> void insertDataToWorkbook(Workbook workbook, ExportConfig exportConfig, List<T> datas,
                                                 XSSFCellStyle dataCellStyle) {
        int startRowIndex = exportConfig.getStartRow();//2

        int sheetIndex = exportConfig.getSheetIndex();//1

        Class clazz = exportConfig.getDataClazz();

        List<CellConfig> cellConfigs = exportConfig.getCellExportConfigList();

        Sheet sheet = workbook.getSheetAt(sheetIndex);

        int currentRowIndex = startRowIndex;

        for (T data : datas) {
            Row currentRow = sheet.getRow(currentRowIndex);
            if (ObjectUtils.isEmpty(currentRow)) {
                currentRow = sheet.createRow(currentRowIndex);
            }
            //insert data to row
            insertDataToCell(data, currentRow, cellConfigs, clazz, sheet, dataCellStyle);
            currentRowIndex++;
        }
    }

    private static <T> void insertFieldNameAsTitleToWorkbook(List<CellConfig> cellConfigs,
                                                             Sheet sheet,
                                                             XSSFCellStyle titleCellStyle) {

        //title -> first row of excel -> get top row
        int currentRow = sheet.getTopRow();

        //create row
        Row row = sheet.createRow(currentRow);
        int i = 0;

        //resize fix text in each cell
        sheet.autoSizeColumn(currentRow);

        //insert field name to cell
        for (CellConfig cellConfig : cellConfigs) {
            Cell currentCell = row.createCell(i);
            String fieldName = cellConfig.getFieldName();
            currentCell.setCellValue(fieldName);
            currentCell.setCellStyle(titleCellStyle);
            sheet.autoSizeColumn(i);
            i++;
        }

    }

    private static <T> void insertDataToCell(T data, Row currentRow, List<CellConfig> cellConfigs,
                                             Class clazz, Sheet sheet, XSSFCellStyle dataStyle) {

        for (CellConfig cellConfig : cellConfigs) {
            Cell currentCell = currentRow.getCell(cellConfig.getColumnIndex());
            if (ObjectUtils.isEmpty(currentCell)) {
                currentCell = currentRow.createCell(cellConfig.getColumnIndex());
            }

            //get data for cell
            String cellValue = getCellValue(data, cellConfig, clazz);

            //set data
            currentCell.setCellValue(cellValue);
            sheet.autoSizeColumn(cellConfig.getColumnIndex());
            currentCell.setCellStyle(dataStyle);
        }

    }

    private static <T> String getCellValue(T data, CellConfig cellConfig, Class clazz) {
        String fieldName = cellConfig.getFieldName();
        try {
            Field field = getDeclaredField(clazz, fieldName);
            if (!ObjectUtils.isEmpty(field)) {
                field.setAccessible(true);
                return !ObjectUtils.isEmpty(field.get(data)) ? field.get(data).toString() : "";
            }
            return "";
        } catch (Exception e) {
            log.info("" + e);
            return "";
        }
    }

    private static Field getDeclaredField(Class clazz, String fieldName) {
        if (ObjectUtils.isEmpty(clazz) || ObjectUtils.isEmpty(fieldName)) {
            return null;
        }
        do {
            try {
                Field field = clazz.getDeclaredField(fieldName);
                field.setAccessible(true);
                return field;
            } catch (Exception e) {
                log.info("" + e);
            }
        } while ((clazz = clazz.getSuperclass()) != null);

        return null;
    }
    //end


    //import config
    public static <T> List<T> getImportData(Workbook workbook, ImportConfig importConfig) {
        List<T> list = new ArrayList<>();

        List<CellConfig> cellConfigs = importConfig.getCellImportConfigs();

        int countSheet = 0;

        for (Sheet sheet : workbook) {
            if (countSheet != importConfig.getSheetIndex()) {
                countSheet++;
                continue;
            }

            int countRow = 0;
            for (Row row : sheet) {
                if (countRow < importConfig.getStartRow()) {
                    countRow++;
                    continue;
                }
                T rowData = getRowData(row, cellConfigs, importConfig.getDataClazz());
                list.add(rowData);
                countRow++;
            }
            countSheet++;
        }
        return list;
    }

    private static <T> T getRowData(Row row, List<CellConfig> cellConfigs, Class dataClazz) {
        T instance = null;
        try {
            instance = (T) dataClazz.getDeclaredConstructor().newInstance();

            for (int i = 0; i < cellConfigs.size(); i++) {
                CellConfig currentCell = cellConfigs.get(i);
                try {
                    Field field = getDeclaredField(dataClazz, currentCell.getFieldName());

                    Cell cell = row.getCell(currentCell.getColumnIndex());
                    if (!ObjectUtils.isEmpty(cell)) {
                        cell.setCellType(CellType.STRING);

                        Object cellValue = cell.getStringCellValue();

                        setFieldValue(instance, field, cellValue);
                    }

                } catch (Exception e) {
                    e.printStackTrace();
                    return null;
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }

        return instance;

    }

    private static <T> void setFieldValue(Object instance, Field field, Object cellValue) {
        if (ObjectUtils.isEmpty(instance) || ObjectUtils.isEmpty(field)) {
            return;
        }

        Class clazz = field.getType();

        Object valueConverted = parseValueByClass(clazz, cellValue);

        field.setAccessible(true);

        try {
            field.set(instance, valueConverted);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static Object parseValueByClass(Class clazz, Object cellValue) {
        if (ObjectUtils.isEmpty(cellValue) || ObjectUtils.isEmpty(clazz)) {
            return null;
        }
        String clazzName = clazz.getSimpleName();

        switch (clazzName) {
            case "char":
                cellValue = parseChar(cellValue);
                break;
            case "String":
                cellValue = cellValue.toString().trim();
                break;
            case "boolean":
            case "Boolean":
                cellValue = parseBoolean(cellValue);
                break;
            case "byte":
            case "Byte":
                cellValue = parseByte(cellValue);
                break;
            case "short":
            case "Short":
                cellValue = parseShort(cellValue);
                break;
            case "int":
            case "Integer":
                cellValue = parseInt(cellValue);
                break;
            case "long":
            case "Long":
                cellValue = parseLong(cellValue);
                break;
            case "float":
            case "Float":
                cellValue = parseFloat(cellValue);
                break;
            case "double":
            case "Double":
                cellValue = parseDouble(cellValue);
                break;
            case "Date":
                cellValue = parseDate(cellValue);
                break;
            case "Instant":
                cellValue = parseInstant(cellValue);
                break;
            case "Enum":
                cellValue = parseEnum(cellValue, clazz);
                break;
            case "Map":
                cellValue = parseMap(cellValue);
                break;
            case "BigDecimal":
                cellValue = parseBigDecimal(cellValue);
                break;
            default:
                break;
        }
        return cellValue;
    }


    private static Object parseChar(Object value) {
        return ObjectUtils.isEmpty(value) ? null : (char) value;
    }


    private static Object parseBoolean(Object value) {
        return ObjectUtils.isEmpty(value) ? null : (Boolean) value;
    }

    private static Object parseMap(Object value) {
        if (ObjectUtils.isEmpty(value)) {
            return null;
        }
        return (Map) value;
    }

    private static Object parseEnum(Object value, Class clazz) {
        if (ObjectUtils.isEmpty(value)) {
            return null;
        }

        String valueStr = value.toString().trim();
        return Enum.valueOf(clazz, valueStr);
    }

    private static Date parseDate(Object value) {
        String[] formatsDate = {"yyyy-MM-dd HH:mm:ss", "dd/MM/yyyy"};

        if (ObjectUtils.isEmpty(value)) {
            return null;
        }

        String dateStr = value.toString();
        for (String format : formatsDate) {
            Date date = null;

            try {
                DateFormat dateFormat = new SimpleDateFormat(format);
                date = dateFormat.parse(dateStr);
            } catch (Exception e) {
                e.printStackTrace();
            }
            if (ObjectUtils.isEmpty(date)) {
                return date;
            }
        }

        try {
            Date date = (Date) value;
            return date;
        } catch (Exception e) {
            e.printStackTrace();
            return new Date();
        }
    }

    private static Object parseInstant(Object value) {
        return ObjectUtils.isEmpty(value) ? null : parseDate(value).toInstant();
    }

    private static Double parseDouble(Object value) {
        if (ObjectUtils.isEmpty(value)) {
            return null;
        }

        try {
            return Double.parseDouble(value.toString());
        } catch (Exception e) {
            return null;
        }
    }

    private static Object parseFloat(Object value) {
        if (ObjectUtils.isEmpty(value)) {
            return null;
        }

        Double doubleValue = parseDouble(value);
        return ObjectUtils.isEmpty(doubleValue) ? null : doubleValue.floatValue();
    }


    private static Object parseLong(Object value) {
        if (ObjectUtils.isEmpty(value)) {
            return null;
        }

        Double longDoubleValue = parseDouble(value);
        return ObjectUtils.isEmpty(longDoubleValue) ? null : longDoubleValue.longValue();
    }

    private static Object parseShort(Object value) {
        if (ObjectUtils.isEmpty(value)) {
            return null;
        }

        Double shortDoubleValue = parseDouble(value);
        return ObjectUtils.isEmpty(shortDoubleValue) ? null : shortDoubleValue.shortValue();
    }

    private static Object parseInt(Object value) {
        if (ObjectUtils.isEmpty(value)) {
            return null;
        }

        Double intDoubleValue = parseDouble(value);
        return ObjectUtils.isEmpty(intDoubleValue) ? null : intDoubleValue.intValue();
    }

    private static Object parseBigDecimal(Object value) {
        if (ObjectUtils.isEmpty(value)) {
            return null;
        }

        try {
            return BigDecimal.valueOf(Double.valueOf(value.toString()));
        } catch (Exception e) {
            return null;
        }
    }

    private static Object parseByte(Object value) {
        if (ObjectUtils.isEmpty(value)) {
            return null;
        }

        Double byteDoubleValue = parseDouble(value);
        return ObjectUtils.isEmpty(byteDoubleValue) ? null : byteDoubleValue.byteValue();
    }


}
