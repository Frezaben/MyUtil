package cqie.per.springtest.util.excel.factorys.writer;

import com.github.xiaoymin.knife4j.annotations.Ignore;
import cqie.per.springtest.util.excel.annotations.ExcelCell;
import cqie.per.springtest.util.excel.exception.ExcelReadException;
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
import java.util.Map;

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
        }finally {
            try {
                inputStream.close();
            } catch (IOException e) {
                log.error("Can not close file ,"+e.getMessage());
            }
        }
        return workbook;
    }

    public <E> Sheet writeRow(Sheet sheet, int rowIndex, E data, Map<String,String> headColumn){
        Row row = sheet.getRow(rowIndex);
        if(null == row){
            row = sheet.createRow(rowIndex);
        }
        Class<?> cls = data.getClass();
        Row finalRow = row;
        headColumn.forEach((key, value)->{
            int celIndex = getCellNum(key);
            try {
                writeCell(cls.getDeclaredField(value),data,celIndex, finalRow);
            } catch (NoSuchFieldException e) {
                log.warn("Fail to write cell, can not find field:'"+value+"'");
            }
        });
        return sheet;
    }

    public <E> Sheet writeRow(Sheet sheet, int rowIndex, E data){
        Row row = sheet.getRow(rowIndex);
        if(null == row){
            row = sheet.createRow(rowIndex);
        }
        Class<?> cls = data.getClass();
        Row finalRow = row;
        Field[] fields = cls.getDeclaredFields();
        for (Field field : fields){
            ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
            if(null == excelCell || excelCell.outputIgnore()){
                log.warn("can not mapping cell, cause:'ExcelCell' annotation not found at field '{}'",field.getName());
                continue;
            }
            String cellStr = excelCell.cell();
            int cellIndex;
            if(excelCell.useIndex()){
                cellIndex = getCellNum(cellStr);
            }else {
                log.warn("can not mapping cell, cause:must set 'useIndex' true at field '{}'",field.getName());
                continue;
            }
            writeCell(field,data,cellIndex,finalRow);
        }
        return sheet;
    }

    public static void writeCell(Field field,Object data, int cellIndex, Row excelRow) {
        field.setAccessible(true);
        Object filedValue;
        try {
            filedValue = field.get(data);
        } catch (IllegalAccessException e) {
            log.warn("Fail to write cell, can not access field:'"+field.getName()+"'");
            return;
        }
        Cell cell = excelRow.createCell(cellIndex);
        Class<?> fieldCls =field.getType();
        if(fieldCls.isInstance("")){
            cell.setCellValue(filedValue.toString());
        }else if(fieldCls.isInstance(0)||fieldCls.isInstance(Double.valueOf("0"))||fieldCls.isInstance(0L)){
            cell.setCellValue(((Number) filedValue).doubleValue());
        } else if(fieldCls.isInstance(new Date())){
            cell.setCellValue((Date) filedValue);
        }
    }

    public int getCellNum(String cell){
        if("".equals(cell) || null == cell){
            throw new ExcelReadException("Unknown index");
        }
        cell = cell.toUpperCase();
        char[] chars = cell.toCharArray();
        int num  = 0;
        for (char str : chars){
            num += str - 'A';
        }
        return num;
    }

    public String getCellStr(int cell){
        int intNum = cell / 26;
        int remNum = cell % 26;
        String str = "";
        str += (char) intNum;
        if(remNum!=0){
            str += (char) remNum;
        }
        return str;
    }
}
