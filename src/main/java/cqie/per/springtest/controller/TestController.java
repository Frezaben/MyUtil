package cqie.per.springtest.controller;

import cqie.per.springtest.entity.StudentInfo;
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


    @PostMapping(value = "xxx", consumes = {"multipart/form-data"})
    @ApiOperation("test")
    public List<StudentInfo> uploadBom(@RequestParam MultipartFile file){
        System.out.println(file.getName());
        return ExcelUtil.read(file, StudentInfo.class);

    }

    @PostMapping(value = "xxxx")
    @ApiOperation("test")
    public List<UserInfo> xxxxx(@RequestParam MultipartFile file){
        System.out.println(file.getName());
        return ExcelUtil.read(file,UserInfo.class);

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
    public String xx() throws IllegalAccessException {
        List<StudentInfo> studentList = new ArrayList<>();
        StudentInfo studentInfo = new StudentInfo();
        studentInfo.setName("xx");
        studentInfo.setUid(18900);
        StudentInfo studentInfo1 = new StudentInfo();
        studentInfo1.setName("xxx");
        studentInfo1.setUid(1800);
        studentList.add(studentInfo);
        studentList.add(studentInfo1);
        Workbook workbook = ExcelUtil.write("C:\\Users\\benye\\Desktop\\tem.xlsx",studentList,StudentInfo.class);
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