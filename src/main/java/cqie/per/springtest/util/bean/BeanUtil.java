package cqie.per.springtest.util.bean;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Arrays;
import java.util.Set;
import java.util.stream.Collectors;

public class BeanUtil {
    /**
     * @param o 数据源 data scours
     * @param cls 目标对象 target class
     * @param ignoreProperty 忽略属性 ignore field
     * @param <T> 返回实体类，根据目标对象返回 return
     * 复制数据源中属性，返回新目标对象
     */
    public static <T> T copy(Object o, Class<T> cls, String ...ignoreProperty){
        T result;
        try {
            result= cls.getDeclaredConstructor().newInstance();
        } catch (InstantiationException | InvocationTargetException | IllegalAccessException e) {
            throw new ConvertException("can not convert"+e.getMessage());
        } catch (NoSuchMethodException e) {
            throw new ConvertException("Instance '"+cls.getName()+"' fail, need No-parameter Constructor");
        }
        Class<?> scours = o.getClass();
        Set<String> fieldNames = Arrays.stream(cls.getDeclaredFields()).map(Field::getName).collect(Collectors.toSet());
        Set<String> intersection = Arrays.stream( scours.getDeclaredFields()).map(Field::getName).collect(Collectors.toSet());
        fieldNames.retainAll(intersection);
        if(null != ignoreProperty && ignoreProperty.length>0){
            Arrays.asList(ignoreProperty).forEach(fieldNames::remove);
        }
        for(String name : fieldNames){
            Field field;
            Method method;
            try {
                field = scours.getDeclaredField(name);
            } catch (NoSuchFieldException e) {
                throw new ConvertException("can not convert, no property name :'"+name+"'found in '"+scours.getName()+"'");
            }
            field.setAccessible(true);
            String setter = "set".concat(field.getName().substring(0,1).toUpperCase())
                    .concat(field.getName().substring(1));
            try {
                method = cls.getMethod(setter,field.getType());
            } catch (NoSuchMethodException e) {
                throw new ConvertException("can not set property on '"+name+"'," +
                        " no public setter found in '"+cls.getName()+"'"+"with arg:'"+field.getType()+"'");
            }
            try {
                method.invoke(result,field.get(o));
            } catch (IllegalAccessException | InvocationTargetException e) {
                throw new ConvertException("can not set property on '"+name+"', "+e.getMessage());
            }
        }
        return result;
    }

    /**
     * @param cls 实例对象 instance target class
     * @param <T> 返回实例 return
     */
    public static <T> T getInstance(Class<T> cls){
        T data;
        try {
            data = cls.getDeclaredConstructor().newInstance();
        } catch (InstantiationException | InvocationTargetException | IllegalAccessException e) {
            throw new ConvertException("can not Instance class :‘"+cls.getName()+"’"+e.getMessage());
        } catch (NoSuchMethodException e) {
            throw new ConvertException("Instance '"+cls.getName()+"' fail, need No-parameter Constructor");
        }
        return data;
    }

}
