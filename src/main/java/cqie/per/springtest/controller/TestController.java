package cqie.per.springtest.controller;

import cqie.per.springtest.entity.UserInfo;
import cqie.per.springtest.util.excel.ExcelUtil;
import io.swagger.annotations.ApiOperation;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

@RestController
public class TestController {

    @PostMapping(value = "xxx")
    @ApiOperation("test")
    public List<UserInfo> uploadBom(@RequestBody MultipartFile file){
        System.out.println(file.getName());
        return ExcelUtil.readExcel(file,UserInfo.class);
    }
}
