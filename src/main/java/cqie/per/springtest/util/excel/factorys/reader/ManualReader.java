package cqie.per.springtest.util.excel.factorys.reader;

import cqie.per.springtest.util.excel.annotations.ExcelCell;
import cqie.per.springtest.util.excel.annotations.ExcelSheet;
import cqie.per.springtest.util.excel.exception.ExcelReadException;
import cqie.per.springtest.util.excel.factorys.ReaderFactory;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.util.StringUtils;
import org.springframework.web.multipart.MultipartFile;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

@Slf4j
public class ManualReader implements ReaderFactory {
    private final CommonReader commonReader = new CommonReader();

    @Override
    public <E> List<E> readExcel(MultipartFile file, Class<E> cls) {
        ExcelSheet annotation = cls.getAnnotation(ExcelSheet.class);
        if(null == annotation){
            throw new ExcelReadException("Can not read excel with manual module ," +
                    " cause: 'ExcelSheet' annotation not found in class:'"+cls.getName()+"'");
        }
        if(!StringUtils.hasText(annotation.sheet())){
            throw new ExcelReadException("Can not read excel with manual module ,"+
                    " cause: sheet name not assign");
        }
        Workbook workbook = commonReader.getWorkbook(file);
        Sheet sheet = workbook.getSheet(annotation.sheet());
        if(null == sheet){
            return Collections.emptyList();
        }
        List<E> dataList = new ArrayList<>();
        int startRow = annotation.dataRow();
        int lastRow = sheet.getLastRowNum();
        for (; startRow <= lastRow; startRow++){
            E data = commonReader.getInstance(cls);
            readRow(sheet.getRow(startRow),data);
            dataList.add(data);
        }
        return dataList;
    }

    @Override
    public <E> List<E> readExcel(Workbook workbook, Class<E> cls) {
        return null;
    }

    private <E> void readRow(Row row, E data) { Class<?> cls = data.getClass();
        Field[] dataFields = cls.getDeclaredFields();
        for(Field field : dataFields){
            int celNum;
            ExcelCell cellAnnotation = field.getAnnotation(ExcelCell.class);
            if(null == cellAnnotation){
                continue;
            }
            celNum = commonReader.getCellNum(cellAnnotation.cell());
            try {
                commonReader.readCell(row, data, field, celNum);
            }catch (IllegalArgumentException e){
                log.error("Invalid column index (" + celNum + "), using manual module must use index at annotation:'Cell()'");
            }
        }
    }
}
