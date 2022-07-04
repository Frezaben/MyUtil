package cqie.per.springtest.util.excel.factorys.writer;

import cqie.per.springtest.util.excel.exception.ExcelReadException;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.HashMap;

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
}
