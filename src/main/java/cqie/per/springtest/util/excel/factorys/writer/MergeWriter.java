package cqie.per.springtest.util.excel.factorys.writer;

import cqie.per.springtest.util.excel.annotations.ExcelSheet;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.lang.reflect.Field;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Objects;

@Slf4j
public class MergeWriter {
    private final CommonWriter commonWriter = new CommonWriter();

    public <E> Workbook write(List<E> dataList, String tempPath, Sheet sheet) throws IllegalAccessException {
        Workbook workbook = commonWriter.getWorkbook(tempPath);
        Class<?> cls = dataList.get(0).getClass();
        ExcelSheet annotation = cls.getAnnotation(ExcelSheet.class);
        String sheetName = annotation.sheet();
        HashMap<Integer, String> headCell =
                commonWriter.getHeadCell(workbook.getSheet(sheetName),annotation.sectionRow());
        int columnSize = headCell.size();
        for(int column = 0; column < columnSize;column++){
            Field field;
            try {
                field = cls.getDeclaredField(headCell.get(column));
            }catch (NoSuchFieldException e){
                log.error("Can not found field:'{}' in temple:{}",headCell.get(column),tempPath);
                continue;
            }
            field.setAccessible(true);

            int dataIndex = 0;
            String startValue = field.get(dataList.get(dataIndex)).toString();
            int startRowIndex = annotation.dataRow();
            String currentValue;

            for (int row = annotation.dataRow(); row <= dataList.size(); row++){
                Row excelRow = sheet.getRow(row);
                if(null == excelRow){
                    excelRow = sheet.createRow(row);
                }
                E outputData = dataList.get(dataIndex++);
                currentValue = field.get(outputData).toString();
                if(Objects.equals(startValue,currentValue) && startRowIndex != row){
                   continue;
                }
                if(row-startRowIndex>1) {
                    CellRangeAddress cellRangeAddress = new CellRangeAddress(startRowIndex, row-1, column, column);
                    sheet.addMergedRegion(cellRangeAddress);
                }
                writeCell(field, outputData, column, excelRow);
                startRowIndex = row;
                startValue = currentValue;
            }
        }
        return workbook;
    }

    private void writeCell(Field field,Object data, int cellIndex, Row excelRow) {
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
