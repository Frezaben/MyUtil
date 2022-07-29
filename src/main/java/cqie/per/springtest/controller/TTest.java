package cqie.per.springtest.controller;

import cqie.per.springtest.entity.UserInfo;
import cqie.per.springtest.util.excel.ExcelUtil;
import cqie.per.springtest.util.excel.factorys.writer.CommonWriter;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class TTest {
    public static void main(String[] args) throws IllegalAccessException, IOException {
        UserInfo userInfo = new UserInfo();
        userInfo.setName("wangwww");
        userInfo.setAge("19");
        UserInfo userInfo2 = new UserInfo();
        userInfo2.setName("axcwcw");
        userInfo2.setAge("19");
        UserInfo userInfo3 = new UserInfo();
        userInfo3.setName("xxx");
        userInfo3.setAge("190");
        List<UserInfo> test = new ArrayList<>();
        test.add(userInfo);
        test.add(userInfo2);
        test.add(userInfo3);
        Workbook workbook = ExcelUtil.write(test,UserInfo.class);
        File file = new File("C:\\Users\\benye\\Desktop\\output.xlsx");
        FileOutputStream fileOutputStream = new FileOutputStream(file);
        workbook.write(fileOutputStream);
        fileOutputStream.flush();
        fileOutputStream.close();

    }
}

