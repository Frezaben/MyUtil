package cqie.per.springtest.controller;

import cqie.per.springtest.entity.UserInfo;
import cqie.per.springtest.util.excel.ExcelUtil;
import io.swagger.annotations.ApiOperation;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@RestController
public class TestController {

    @PostMapping(value = "xxx")
    @ApiOperation("test")
    public List<UserInfo> uploadBom(@RequestBody MultipartFile file){
        System.out.println(file.getName());
        return ExcelUtil.readExcel(file,UserInfo.class);

    }

    @PostMapping(value = "xxxx")
    @ApiOperation("test")
    public List<UserInfo> xxxxx(@RequestBody MultipartFile file){
        System.out.println(file.getName());
        return ExcelUtil.readExcel(file,UserInfo.class);

    }

    @PostMapping(value = "xx")
    @ApiOperation("byte")
    public byte[] xxxx(@RequestBody MultipartFile file){
        try {
            return file.getBytes();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    @RequestMapping(value = "x")
    @ApiOperation("output")
    public String xx(){
        List<UserInfo> userInfos = new ArrayList<>();
        UserInfo userInfo = new UserInfo();
        userInfo.setAge("10");
        userInfo.setId(10000L);
        userInfo.setUid(101);
        userInfo.setName("wang.de.fa");
        userInfos.add(userInfo);
        UserInfo userInfo1 = new UserInfo();
        userInfo1.setAge("12");
        userInfo1.setId(10000L);
        userInfo1.setUid(101);
        userInfo1.setName("de.de.fa");
        userInfos.add(userInfo1);
        Workbook workbook = ExcelUtil.output("C:\\Users\\benye\\Desktop\\tem.xlsx",userInfos,UserInfo.class);
        FileOutputStream fileOutputStream = null;
        File file = new File("C:\\Users\\benye\\Desktop\\output.xlsx");
        try {
            fileOutputStream = new FileOutputStream(file);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        try {
            workbook.write(fileOutputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        try {
            assert fileOutputStream != null;
            fileOutputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return "ok";

    }
}