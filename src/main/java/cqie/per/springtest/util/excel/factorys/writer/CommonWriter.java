package cqie.per.springtest.util.excel.factorys.writer;

import cqie.per.springtest.util.excel.annotations.ExcelCell;
import cqie.per.springtest.util.excel.annotations.ExcelSheet;
import cqie.per.springtest.util.excel.annotations.OutputIgnore;
import cqie.per.springtest.util.excel.exception.ExcelReadException;
import cqie.per.springtest.util.excel.factorys.reader.CommonReader;
import io.swagger.annotations.ApiModelProperty;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.StringUtils;

import java.io.*;
import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;

@Slf4j
public class CommonWriter {

    public static HashMap<Integer,String> getHeadCell(Sheet sheet, int section){
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

    public static Workbook getWorkbook(String path){
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

    public static <E> Workbook getWorkbook(Class<E> cls){
        Workbook workbook = new XSSFWorkbook();
        ExcelSheet annotation = cls.getAnnotation(ExcelSheet.class);
        Sheet sheet;
        if(null == annotation || !StringUtils.hasText(annotation.sheet())) {
           sheet = workbook.createSheet("sheet1");
        }else {
           sheet = workbook.createSheet(annotation.sheet());
        }
        Field[] fields = cls.getDeclaredFields();

        int sectionRow = 0;
        if(null != annotation){
            sectionRow = annotation.sectionRow();
        }
        List<Field> fieldList = Arrays.stream(fields)
                .filter(item-> null == item.getAnnotation(OutputIgnore.class))
                .collect(Collectors.toList());
        Map<Integer,String> cellHead = new HashMap<>(fieldList.size());
        fieldList.forEach(field -> {
            ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
            if(null == excelCell || !excelCell.useIndex()){
                return;
            }
            String outputName = excelCell.outputName();
            if("".equals(outputName)){
                ApiModelProperty apiModelProperty = field.getAnnotation(ApiModelProperty.class);
                outputName = apiModelProperty.value();
                if("".equals(outputName)){
                    outputName = field.getName();
                }
            }
            cellHead.put(CommonReader.getCellNum(excelCell.cell()),outputName);
        });
        fieldList.removeIf(field ->field.getAnnotation(ExcelCell.class).useIndex());
        Row row = sheet.createRow(sectionRow);
        int index=0;
        for(int column = 0;column< fields.length; column++){
            Cell cell = row.createCell(column);
            if(cellHead.containsKey(column)){
                cell.setCellValue(cellHead.get(column));
                continue;
            }
            cell.setCellValue(fieldList.get(index++).getName());
        }
        return workbook;
    }

    public static <E> void writeColumn(Collection<E> collection, Sheet sheet, int column ,int startRow ,Field field){
        List<E> dataList = collection.stream().toList();
        for (int row = startRow; row <= dataList.size(); row++) {
            Row excelRow = sheet.getRow(row);
            if (null == excelRow) {
                excelRow = sheet.createRow(row);
            }
            E data =  dataList.get(row);
            writeCell(field,data,column,excelRow);
        }
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

}
