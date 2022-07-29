package cqie.per.springtest.util.excel.factorys.writer;

import cqie.per.springtest.util.excel.annotations.ExcelCell;
import cqie.per.springtest.util.excel.annotations.ExcelSheet;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.lang.reflect.Field;
import java.util.Collection;
import java.util.HashMap;
import java.util.List;
import java.util.Objects;

@Slf4j
public class MergeWriter {


    public <E> Workbook write(List<E> dataList, Sheet sheet, String ...ignoreFields) throws IllegalAccessException {
        //复制输出workbook
        Workbook workbook = sheet.getWorkbook();
        //获取类型
        Class<?> cls = dataList.get(0).getClass();
        //获取注解
        ExcelSheet annotation = cls.getAnnotation(ExcelSheet.class);
        String sheetName = annotation.sheet();
        //获取模板的头
        HashMap<Integer, String> headCell =
                CommonWriter.getHeadCell(workbook.getSheet(sheetName),annotation.sectionRow());
        if(null!=ignoreFields&&ignoreFields.length>0){
            Collection<String> values = headCell.values();
            for (String ignoreFiled : ignoreFields){
                values.remove(ignoreFiled);
            }
        }
        int columnSize = headCell.size();
        //外层循环列，遍历列再遍历行，实现相邻格（列）的合并
        for(int column = 0; column < columnSize;column++){
            Field field;
            try {
                //从模板头Map里获取当前列的属性
                field = cls.getDeclaredField(headCell.get(column));
            }catch (NoSuchFieldException e){
                log.error("Can not found field:'{}' in class:{}",headCell.get(column),cls.getName());
                continue;
            }
            //开放属性的访问
            field.setAccessible(true);
            //获取属性的注解
            ExcelCell cellAnnotation = field.getAnnotation(ExcelCell.class);
            //如果不需要合并单元格，则调用CommonWriter方法
            if(null != cellAnnotation && !cellAnnotation.mergedSameRegion()){
                //写入整列
                CommonWriter.writeColumn(dataList,sheet,column,annotation.dataRow(),field);
                //跳过后续，开始下一列
                continue;
            }

            //记录当前列开始行的位置，与单元格内容
            int dataIndex = 0;
            String startValue = field.get(dataList.get(dataIndex)).toString();
            int startRowIndex = annotation.dataRow();
            String currentValue;

            //内层循环当前列的行
            for (int row = annotation.dataRow(); row <= dataList.size(); row++){
                //获取Row
                Row excelRow = sheet.getRow(row);
                if(null == excelRow){
                    //为空的话新建Row
                    excelRow = sheet.createRow(row);
                }
                //获取当前行在数据列表dataList的对象
                E outputData = dataList.get(dataIndex++);
                //转换为String方便判断是否相等
                currentValue = field.get(outputData).toString();
                //如果当前行与开始行的值相等，先跳过不插入（等出现不同值后才将之前的合并插入）
                if(Objects.equals(startValue,currentValue) && startRowIndex != row){
                   continue;
                }
                //与当前行不同

                //如果只有一行不同不需要进行单元格合并，如果大于1行则进行单元格合并
                if(row-startRowIndex>1) {
                    //初始行到当前行单元格合并
                    CellRangeAddress cellRangeAddress = new CellRangeAddress(startRowIndex, row-1, column, column);
                    sheet.addMergedRegion(cellRangeAddress);
                }
                //写入值（注：合并的单元格需要把值插入到开始行处）
                CommonWriter.writeCell(field, outputData, column, excelRow);
                //修改开始行为当前行
                startRowIndex = row;
                //修改开始值为当前行的值
                startValue = currentValue;
            }
        }
        //复制输出workbook
        return sheet.getWorkbook();
    }



}
