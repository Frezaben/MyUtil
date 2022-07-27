package cqie.per.springtest.util.excel;

import cqie.per.springtest.util.excel.annotations.ExcelSheet;
import cqie.per.springtest.util.excel.exception.ExcelReadException;
import cqie.per.springtest.util.excel.factorys.ReaderFactory;
import cqie.per.springtest.util.excel.factorys.reader.AutoReader;
import cqie.per.springtest.util.excel.factorys.reader.DefaultReader;
import cqie.per.springtest.util.excel.factorys.reader.ManualReader;
import cqie.per.springtest.util.excel.factorys.writer.CommonWriter;
import cqie.per.springtest.util.excel.factorys.writer.MergeWriter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.springframework.web.multipart.MultipartFile;

import java.util.*;

@Slf4j
public class ExcelUtil {

    public static <E> List<E> read(MultipartFile file,Class<E> cls){
        ExcelSheet annotation =  cls.getAnnotation(ExcelSheet.class);
        ReaderFactory reader;
        if(null == annotation){
            log.warn("'ExcelSheet annotation not found in class:'"+cls.getName()+"', run with default Reader'");
            reader = new DefaultReader();
            return reader.readExcel(file, cls);
        }
        switch (annotation.model()){
            case AUTO -> reader = new AutoReader();
            case MANUAL -> reader = new ManualReader();
            case DEFAULT -> reader = new DefaultReader();
            default -> throw new ExcelReadException("failed to Instantiation 'Reader', unknown model:'"+annotation.model()+"'");
        }
        return reader.readExcel(file,cls);
    }

    public static <E> List<E> read(MultipartFile file,Class<E> cls,ReaderFactory reader){
        return reader.readExcel(file,cls);
    }

    public static <E> List<E> read(Workbook workbook,Class<E> cls,ReaderFactory reader){
        return reader.readExcel(workbook,cls);
    }

    public static <E> Workbook write(String templatePaths, Collection<E> collection, Class<E> cls, String ...ignoreFields) throws IllegalAccessException {
        CommonWriter writer = new CommonWriter();
        Workbook template = writer.getWorkbook(templatePaths);
        if (null == collection || collection.size()==0){
            return null;
        }
        List<E> dataList = collection.stream().toList();
        ExcelSheet annotation = cls.getAnnotation(ExcelSheet.class);
        Sheet sheet;
        String sheetName = annotation.sheet();
        if("".equals(sheetName)){
            sheet = template.getSheetAt(0);
        }else {
            sheet = template.getSheet(sheetName);
        }
        HashMap<Integer,String> headRow = writer.getHeadCell(sheet,annotation.sectionRow());
        if(null!=ignoreFields&&ignoreFields.length>0){
            Collection<String> values = headRow.values();
            for (String ignoreFiled : ignoreFields){
                values.remove(ignoreFiled);
            }
        }
        MergeWriter mergeWriter = new MergeWriter();
        return mergeWriter.write(dataList,templatePaths,sheet);
    }
 }
