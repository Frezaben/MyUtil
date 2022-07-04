package cqie.per.springtest.util.excel.factorys;

import org.springframework.web.multipart.MultipartFile;

import java.util.List;

public interface ReaderFactory {
    <E> List<E> readExcel(MultipartFile file, Class<E> cls);
}
