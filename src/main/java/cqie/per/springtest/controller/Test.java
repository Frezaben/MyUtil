package cqie.per.springtest.controller;


import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;

public class Test {
    public static void main(String[] args) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
//        JsonTest jsonTest = new JsonTest();
//        Map<Integer,String> map2 = new HashMap<>();
//        map2.put(1,"diyi");
//        map2.put(2,"dier");
//        Map<String,Integer> map1 = new HashMap<>();
//        String dawd = """
//                dddd\n\r\0
//                .x2222
//                """;
//        map1.put(dawd,10);
//        map1.put("tow",20);
//        jsonTest.setMap1(map1);
//        jsonTest.setMap2(map2);
//        List<String> list = new ArrayList<>();
//        list.add("xxx");
//        list.add("xxxx");
//        jsonTest.setListString(list);
//        jsonTest.setNumBigD(new BigDecimal(100));
//        jsonTest.setNumInt(9);
//        jsonTest.setNumLon(90L);
//        jsonTest.setStr("ssssssssssss");
//        System.out.println(BeanUtil.toJSONString(jsonTest));
//        try {
//            FileInputStream fileInputStream = new FileInputStream("");
//            BufferedInputStream bufferedInputStream = new BufferedInputStream(fileInputStream);
//            FileOutputStream fileOutputStream = new FileOutputStream("");
//            BufferedOutputStream bufferedOutputStream = new BufferedOutputStream(fileOutputStream);
//            byte[] bytes = new byte[1024];
//            while (-1!=bufferedInputStream.read(bytes,0,bytes.length)){
//                bufferedOutputStream.write(bytes);
//                bufferedOutputStream.flush();
//            }
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
        List<Double> strings = new ArrayList<>();
        strings.add(123.9902);
        strings.getClass().getMethod("add",Object.class).invoke(strings,"String");
        System.out.println(strings);

        List list = new ArrayList<>();
        list.add(123.110);
        list.add("str");
        test(list);

    }
    public static void test(List<String> list){
        list.forEach(item->{
            item.concat("x");
        });
    }

}
