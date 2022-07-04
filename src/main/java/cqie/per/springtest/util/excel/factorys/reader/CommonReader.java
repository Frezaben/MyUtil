package cqie.per.springtest.util.excel.factorys.reader;

import cqie.per.springtest.util.bean.ConvertException;
import cqie.per.springtest.util.excel.annotations.ExcelCell;
import cqie.per.springtest.util.excel.exception.ExcelReadException;
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
import java.util.Date;
import java.util.HashMap;

@Slf4j
public class CommonReader {

    public Workbook getWorkbook(MultipartFile file){
        Workbook workbook = null;
        String fileName = file.getOriginalFilename();
        if (null == fileName){
            throw new ExcelReadException("Can not read file ,cause: empty file name ");
        }
        InputStream inputStream = null;
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
        } catch (Exception e){
           log.error(e.getMessage());
        }  finally{
            if(null != inputStream) {
                try {
                    inputStream.close();
                } catch (IOException e) {
                    log.error(e.getMessage());
                }
            }
        }
        return workbook;
    }

    public HashMap<String,Integer> getHeadCell(Sheet sheet, int section){
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

    public <E> void readCell(Row row, E data, Field field, int celNum){
        Cell cell = row.getCell(celNum);
        if(null == cell){
            return;
        }
        ExcelCell annotation = field.getAnnotation(ExcelCell.class);
        boolean nullable = true;
        if(null != annotation){
            if (annotation.readIgnore()){
                return;
            }
            nullable = annotation.outputNotNull();
        }
        if(cell.getCellType() == CellType.BLANK ){
            if (nullable){
                return;
            }
            throw new ExcelReadException("Cell(row:"+row.getRowNum()+",column:"+getCellStr(celNum)+") is empty");
        }
        Class<?> type = field.getType();
        DataFormatter dataFormatter = new DataFormatter();
        field.setAccessible(true);
        try {
            if (type.isInstance("")) {
                field.set(data, dataFormatter.formatCellValue(cell));
            } else if (type.isInstance(0L)) {
                field.set(data, (long) cell.getNumericCellValue());
            } else if (type.isInstance(0)) {
                field.set(data, (int) cell.getNumericCellValue());
            } else if (type.isInstance(new BigDecimal(0))) {
                field.set(data, BigDecimal.valueOf(cell.getNumericCellValue()));
            } else if (type.isInstance((double) 0)) {
                field.set(data, cell.getNumericCellValue());
            } else if (type.isInstance(new Date())) {
                field.set(data, cell.getDateCellValue());
            } else if (type.isInstance(true)) {
                field.set(data, cell.getBooleanCellValue());
            }
        } catch (IllegalAccessException e) {
            throw new ExcelReadException("can not access field:'"+field.getName()+"'"+e.getMessage());
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

    public <E> E getInstance(Class<E> cls){
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
