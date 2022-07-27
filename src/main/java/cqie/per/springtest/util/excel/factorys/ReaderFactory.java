package cqie.per.springtest.util.excel.factorys;

import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

public interface ReaderFactory {
    <E> List<E> readExcel(MultipartFile file, Class<E> cls);
    <E> List<E> readExcel(Workbook workbook, Class<E> cls);
}
