package cqie.per.springtest.util.excel;

import cqie.per.springtest.util.bean.ConvertException;
import io.swagger.annotations.ApiModelProperty;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;

@Slf4j
public class ExcelUtil {
    private static final int SECTION = 0;
    private static final String SHEET_NAME = "Sheet1";

    /**
     * 读Excel（支持xlsx与xls），返回对应数据类列表
     * @param file 文件 excel file
     * @param cls 目标类 target class
     */
    public static <E> List<E> readExcel(MultipartFile file,Class<E> cls){
        Workbook workbook = getWorkbook(file);
        ExcelSheet annotation = cls.getAnnotation(ExcelSheet.class);
        List<E> dataList;
        Sheet sheet;
        if(null == annotation){
            log.warn("ExcelSheet annotation not found in '"+cls.getName()+"' , run with default configuration");
            sheet = workbook.getSheet(SHEET_NAME);
        }else {
            String sheetName = annotation.sheet();
            if("".equals(sheetName)){
                sheetName = SHEET_NAME;
                log.warn("Sheet name not found in '"+cls.getName()+"' , run with default sheet name:'"+SHEET_NAME+"'");
            }
            sheet = workbook.getSheet(sheetName);
        }
        dataList = readSheet(sheet,cls,annotation);
        return dataList;
    }

    private static <E> List<E> readSheet(Sheet sheet,Class<E> cls, ExcelSheet annotation){
        int startRow = 1;
        int lastRow = sheet.getLastRowNum();
        List<E> dataList = new ArrayList<>();
        if(null != annotation) {
            startRow = annotation.data();
            HashMap<String,Integer> headRow = getCellHead(sheet,annotation.section());
            if(annotation.auto()) {
                for (; startRow <= lastRow; startRow++) {
                    E data = getInstance(cls);
                    readRowAuto(sheet.getRow(startRow), data,headRow);
                    dataList.add(data);
                }
            }else {
                for (; startRow <= lastRow; startRow++) {
                    E data = getInstance(cls);
                    readRow(sheet.getRow(startRow), data,headRow);
                    dataList.add(data);
                }
            }
        }else {
            HashMap<String,Integer> headRow = getCellHead(sheet,SECTION);
            for (;startRow<=lastRow;startRow++){
                E data = getInstance(cls);
                readRowAuto(sheet.getRow(startRow),data,headRow);
                dataList.add(data);
            }
        }

        return dataList;
    }

    private static <E> void readRowAuto(Row row, E data,HashMap<String,Integer> headRow) {
        Class<?> cls = data.getClass();
        Field[] dataFields = cls.getDeclaredFields();
        for(Field field : dataFields){
            Integer celNum = null;
            ApiModelProperty annotation = field.getAnnotation(ApiModelProperty.class);
            ExcelCell cellAnnotation = field.getAnnotation(ExcelCell.class);
            if(null != cellAnnotation){
                celNum = headRow.get(cellAnnotation.cell());
            }
            if(null != annotation) {
                celNum = headRow.get(annotation.value());
            }
            if(null == celNum) {
                celNum = headRow.get(field.getName());
            }
            if (null == celNum){
                continue;
            }
            try {
                readCell(row, data, field, celNum,cellAnnotation);
            } catch (IllegalAccessException e) {
                throw new ExcelReadException("can not access field:'"+field.getName()+"'"+e.getMessage());
            }
        }
    }

    private static <E> void readRow(Row row, E data,HashMap<String,Integer> headRow) {
        Class<?> cls = data.getClass();
        Field[] dataFields = cls.getDeclaredFields();
        for(Field field : dataFields){
            ExcelCell cellAnnotation = field.getAnnotation(ExcelCell.class);
            field.setAccessible(true);
            if(null == cellAnnotation){
                continue;
            }
            Integer celNum = headRow.get(cellAnnotation.cell());
            if(null == celNum){
                log.warn("Can not find field:'"+cellAnnotation.cell()+"' in excel, continue processing excel");
                continue;
            }
            try {
                readCell(row, data, field, celNum,cellAnnotation);
            } catch (IllegalAccessException e) {
                throw new ExcelReadException("can not access field:'"+field.getName()+"'"+e.getMessage());
            }
        }
    }

    /**
     * @param row 行 row
     * @param data 数据 data
     * @param field 字段 field
     * @param celNum 目标行 target row
     * @param <E> 数据类型 data type
     */
    private static <E> void readCell(Row row, E data, Field field, int celNum,ExcelCell annotation) throws IllegalAccessException {
        Cell cell = row.getCell(celNum);
        if(null == cell){
            if(annotation.nullable()) {
                field.set(data, null);
                return;
            }
            throw new ExcelReadException("Can not read cell, cause value of cell is NULL. Consider set nullable ture");
        }
        Class<?> type = field.getType();
        DataFormatter dataFormatter = new DataFormatter();
        field.setAccessible(true);
        if (type.isInstance("")){//String
            field.set(data,dataFormatter.formatCellValue(cell));
        }else if(type.isInstance(0L)){//long
            field.set(data,(long)cell.getNumericCellValue());
        } else if (type.isInstance(0)){//int
            field.set(data,(int) cell.getNumericCellValue());
        } else if (type.isInstance(new BigDecimal(0))){//BigDecimal
            field.set(data, BigDecimal.valueOf(cell.getNumericCellValue()));
        } else if (type.isInstance((double) 0)){//double
            field.set(data,cell.getNumericCellValue());
        } else if (type.isInstance(new Date())){//date
            field.set(data,cell.getDateCellValue());
        } else if (type.isInstance(true)){
            field.set(data,cell.getBooleanCellValue());
        }
    }

    /**
     * @param sheet 工作簿 sheet
     * @param section 标题行 section row
     */
    public static HashMap<String,Integer> getCellHead(Sheet sheet,int section){
        Row headRow = sheet.getRow(section);
        HashMap<String,Integer> headCell = new HashMap<>();
        int startCell = headRow.getFirstCellNum();
        int lastCell = headRow.getLastCellNum();
        for(;startCell<lastCell;startCell++){
            Cell cell = headRow.getCell(startCell);
            String name = cell.getStringCellValue();
            if(null == name){
                continue;
            }
            headCell.put(cell.getStringCellValue(),startCell);
        }
        return headCell;
    }

    /**
     * @param file excel文件 excel file
     */
    private static Workbook getWorkbook(MultipartFile file){
        Workbook workbook;
        String fileName = file.getOriginalFilename();
        if (null == fileName){
            throw new ExcelReadException("Can not read file ,cause: empty file name ");
        }
        InputStream inputStream;
        try {
            inputStream= file.getInputStream();
            if(fileName.endsWith("xls")){
                workbook = new HSSFWorkbook(inputStream);
            }else if(fileName.endsWith("xlsx")){
                workbook = new XSSFWorkbook(inputStream);
            }else {
                throw new ExcelReadException("Can not read file '"+fileName+"',unknown file type");
            }
        } catch (IOException e) {
            throw new ExcelReadException("Can not open file ,"+e.getMessage());
        }
        try {
            inputStream.close();
        } catch (IOException e) {
            log.error(e.getMessage());
        }
        return workbook;
    }

    /**
     * @param cls 实例对象
     * @param <E> 实例对象
     */
    private static <E> E getInstance(Class<E> cls){
        E data;
        try {
            data = cls.getDeclaredConstructor().newInstance();
        } catch (InstantiationException | InvocationTargetException | IllegalAccessException e) {
            throw new ConvertException("can not Instance class :‘"+cls.getName()+"’"+e.getMessage());
        } catch (NoSuchMethodException e) {
            throw new ConvertException("Instance '"+cls.getName()+"' fail, need No-parameter Constructor");
        }
        return data;
    }
}
