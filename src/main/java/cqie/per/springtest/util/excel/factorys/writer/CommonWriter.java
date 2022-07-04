package cqie.per.springtest.util.excel.factorys.writer;

import cqie.per.springtest.util.excel.annotations.ExcelCell;
import cqie.per.springtest.util.excel.exception.ExcelReadException;
import io.swagger.annotations.ApiModelProperty;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.lang.reflect.Field;
import java.util.Date;
import java.util.HashMap;

@Slf4j
public class CommonWriter {

    public HashMap<Integer,String> getHeadCell(Sheet sheet, int section){
        Row headRow = sheet.getRow(section);
        HashMap<Integer,String> headCell = new HashMap<>();
        int startCell = headRow.getFirstCellNum();
        int lastCell = headRow.getLastCellNum();
        for(;startCell<lastCell;startCell++){
            Cell cell = headRow.getCell(startCell);
            String name = cell.getStringCellValue();
            if(null == name){
                continue;
            }
            headCell.put(startCell,cell.getStringCellValue());
        }
        return headCell;
    }

    public Workbook getWorkbook(String path){
        Workbook workbook;
        File file = new File(path);
        InputStream inputStream;
        try {
            inputStream = new FileInputStream(path);
        } catch (FileNotFoundException e) {
            throw new ExcelReadException("Can not read file '"+path+"',file not found");
        }
        String fileName = file.getName();
        try {
            if (fileName.endsWith("xls")) {
                workbook = new HSSFWorkbook(inputStream);
            } else if (fileName.endsWith("xlsx")) {
                workbook = new XSSFWorkbook(inputStream);
            } else {
                throw new ExcelReadException("Can not read file '" + fileName + "',unknown file type");
            }
        }catch (IOException e){
            throw new ExcelReadException("Can not open file ,"+e.getMessage());
        }
        return workbook;
    }

    public <E> void writeRow(Row row, E data, HashMap<String,Integer> headRow){
        Class<?> cls = data.getClass();
        headRow.forEach((key,value)->{
            Cell cell = row.createCell(value);
            Field field;
            try {
                field = cls.getDeclaredField(key);
            } catch (NoSuchFieldException e) {
                log.warn("filed :'"+key+"' not found in class:'"+cls.getName()+"'");
                return;
            }
            ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
            if(null == excelCell || excelCell.outputIgnore()){
                return;
            }
            setCellValue(cell,field,data);
        });
    }

    private <E> void writeRowAuto(Row row,E data,HashMap<String,Integer> headRow){
        Class<?> cls = data.getClass();
        Field[] fields = cls.getDeclaredFields();
        for (Field field: fields){
            ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
            if(null != excelCell && excelCell.outputIgnore()){
                return;
            }
            Integer index = null;
            ApiModelProperty annotation = field.getAnnotation(ApiModelProperty.class);
            if(null != annotation){
                index = headRow.get(annotation.value());
            }
            if(null == index){

                if(null != excelCell) {
                    index = headRow.get(excelCell.cell());
                }else {
                    index = headRow.get(field.getName());
                }
            }
            if(null == index){
                log.warn("filed :'"+field.getName()+"' not found in class:'"+cls.getName()+"'");
                return;
            }
            Cell cell = row.createCell(index);
            setCellValue(cell,field,data);
        }
    }

    public <E> void setCellValue(Cell cell,Field field,E data){
        field.setAccessible(true);
        Object filedValue;
        try {
            filedValue = field.get(data);
        } catch (IllegalAccessException e) {
            log.warn("Fail to write cell, can not access field:'"+field.getName()+"'");
            return;
        }
        if(null == filedValue){
            return;
        }
        Class<?> fieldCls =field.getType();
        if(fieldCls.isInstance("")){
            cell.setCellValue(filedValue.toString());
        }else if(fieldCls.isInstance(0)||fieldCls.isInstance(Double.valueOf("0"))||fieldCls.isInstance(0L)){
            cell.setCellValue(((Number) filedValue).doubleValue());
        } else if(fieldCls.isInstance(new Date())){
            cell.setCellValue((Date) filedValue);
        }
    }
}
