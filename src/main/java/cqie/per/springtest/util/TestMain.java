package cqie.per.springtest.util;

import cqie.per.springtest.entity.UserInfo;
import cqie.per.springtest.util.bean.BeanUtil;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

public class TestMain {
    public static void main(String[] args)  {
        UserInfo userInfo = new UserInfo();
        userInfo.setName("王德法");
        userInfo.setUid(100001);
        List<String> list = new ArrayList<>();
        list.add("01");
        list.add("02");
        userInfo.setList(list);
        Set<Integer> set = new HashSet<>();
        set.add(1);
        set.add(2);
        userInfo.setSet(set);
        System.out.println(BeanUtil.toJSONString(userInfo));
    }
}
