package cqie.per.springtest.util.excel.factorys.reader;

import cqie.per.springtest.util.excel.annotations.ExcelCell;
import cqie.per.springtest.util.excel.annotations.ExcelSheet;
import cqie.per.springtest.util.excel.exception.ExcelReadException;
import cqie.per.springtest.util.excel.factorys.ReaderFactory;
import io.swagger.annotations.ApiModelProperty;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.util.StringUtils;
import org.springframework.web.multipart.MultipartFile;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Slf4j
public class AutoReader implements ReaderFactory {
    private final CommonReader commonReader = new CommonReader();
    private static final int SHEET_INDEX = 0;

    @Override
    public <E> List<E> readExcel(MultipartFile file, Class<E> cls) {
        Workbook workbook = commonReader.getWorkbook(file);
        return doRead(workbook,cls);
    }

    @Override
    public <E> List<E> readExcel(Workbook workbook, Class<E> cls) {
       return doRead(workbook,cls);
    }

    private <E> void readRow(Row row, E data, Map<String, Integer> headCell) {
        Class<?> cls = data.getClass();
        Field[] dataFields = cls.getDeclaredFields();
        for(Field field : dataFields){
            Integer celNum;
            ApiModelProperty annotation = field.getAnnotation(ApiModelProperty.class);
            ExcelCell cellAnnotation = field.getAnnotation(ExcelCell.class);
            try {
                if (cellAnnotation.useIndex()) {
                    celNum = commonReader.getCellNum(cellAnnotation.cell());
                } else {
                    celNum = headCell.get(cellAnnotation.cell());
                    if (null != annotation) {
                        celNum = headCell.get(annotation.value());
                    }
                }
                if (null == celNum) {
                    continue;
                }
            }catch (NullPointerException nullPointerException){
                celNum = headCell.get(field.getName());
            }
            commonReader.readCell(row, data, field, celNum);
        }
    }
    private <E> List<E> doRead(Workbook workbook, Class<E> cls){
        ExcelSheet annotation = cls.getAnnotation(ExcelSheet.class);
        if (null == annotation) {
            throw new ExcelReadException("Can not read excel with auto module ," +
                    " cause: 'ExcelSheet' annotation not found in class:'" + cls.getName() + "'");
        }
        Sheet sheet;
        if (StringUtils.hasText(annotation.sheet())) {
            sheet = workbook.getSheet(annotation.sheet());
        } else {
            log.warn("sheet not assign, run with first sheet");
            sheet = workbook.getSheetAt(SHEET_INDEX);
        }
        Map<String, Integer> headCell = commonReader.getHeadCell(sheet, annotation.sectionRow());
        List<E> dataList = new ArrayList<>();
        int startRow = annotation.dataRow();
        int lastRow = sheet.getLastRowNum();
        for (; startRow < lastRow; startRow++) {
            E data = commonReader.getInstance(cls);
            readRow(sheet.getRow(startRow), data, headCell);
            dataList.add(data);
        }
        return dataList;
    }

}
