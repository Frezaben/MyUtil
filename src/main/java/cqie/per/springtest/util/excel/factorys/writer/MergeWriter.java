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

    public <E> Workbook write(List<E> dataList,String tempPath) throws IllegalAccessException {
        Workbook workbook = commonWriter.getWorkbook(tempPath);
        Class<?> cls = dataList.get(0).getClass();
        ExcelSheet annotation = cls.getAnnotation(ExcelSheet.class);
        HashMap<Integer, String> headCell =
                commonWriter.getHeadCell(workbook.getSheet("工作表1"),annotation.sectionRow());
        int columnSize = headCell.size();
        Sheet sheet = workbook.getSheet("工作表1");
        for(int column = 0; column <= columnSize;column++){
            Field field;
            try {
                field = cls.getDeclaredField(headCell.get(column));
            }catch (NoSuchFieldException e){
                log.error("");
                continue;
            }
            field.setAccessible(true);
            int dataIndex = 1;
            String startValue = field.get(dataList.get(dataIndex)).toString();
            int startColumn = 0;
            String currentValue = "";

            for (int row = annotation.dataRow(); row <= dataList.size(); row++){
                Row excelRow = sheet.getRow(row);
                if(null == excelRow){
                    excelRow = sheet.createRow(row);
                }
                E data = dataList.get(dataIndex);
                currentValue = field.get(data).toString();
                if(Objects.equals(startValue,currentValue)){
                    CellRangeAddress cellRangeAddress = new CellRangeAddress(startColumn, row, column, column);
                    sheet.addMergedRegion(cellRangeAddress);
                    writeAndMergeCell(field,data,startColumn,column,row,excelRow);
                    startColumn = column;
                    startValue = currentValue;
                }
            }
        }
        return null;
    }

    private void writeAndMergeCell(Field field,Object data, int startColumn, int column, int row, Row excelRow) {
        field.setAccessible(true);
        Object filedValue;
        try {
            filedValue = field.get(data);
        } catch (IllegalAccessException e) {
            log.warn("Fail to write cell, can not access field:'"+field.getName()+"'");
            return;
        }
        Cell cell = excelRow.createCell(startColumn);
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
