package cqie.per.springtest.util.excel.factorys.reader;

import cqie.per.springtest.util.excel.factorys.ReaderFactory;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.web.multipart.MultipartFile;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Slf4j
public class DefaultReader implements ReaderFactory {
    private final CommonReader commonReader = new CommonReader();
    private static final int SECTION = 1;
    private static final int SHEET_INDEX = 0;

    @Override
    public <E> List<E> readExcel(MultipartFile file, Class<E> cls) {
        Workbook workbook = commonReader.getWorkbook(file);
        Sheet sheet = workbook.getSheetAt(SHEET_INDEX);
        Map<String,Integer> headCell = commonReader.getHeadCell(sheet,SECTION);
        int startRow = SECTION + 1;
        int lastRow = sheet.getLastRowNum();
        List<E> dataList = new ArrayList<>();
        for (; startRow < lastRow; startRow++){
            E data = commonReader.getInstance(cls);
            readRow(sheet.getRow(startRow),data,headCell);
            dataList.add(data);
        }
        return dataList;
    }

    private <E> void readRow(Row row, E data, Map<String,Integer> headRow) {
        Class<?> cls = data.getClass();
        Field[] dataFields = cls.getDeclaredFields();
        for(Field field : dataFields){
            field.setAccessible(true);
            Integer celNum = headRow.get(field.getName());
            if(null == celNum){
                log.warn("Can not find field:'"+field.getName()+"' in excel, continue processing excel");
                continue;
            }
            commonReader.readCell(row, data, field, celNum);
        }
    }

}
