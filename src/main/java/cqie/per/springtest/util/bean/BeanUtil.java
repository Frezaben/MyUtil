package cqie.per.springtest.util.bean;

import lombok.extern.slf4j.Slf4j;

import java.lang.reflect.*;
import java.util.*;
import java.util.stream.Collectors;

@Slf4j
public class BeanUtil {
    /**
     * @param o 数据源 data scours
     * @param cls 目标对象 target class
     * @param ignoreProperty 忽略属性 ignore field
     * @param <T> 返回实体类，根据目标对象返回 return
     * 复制数据源中属性，返回新目标对象
     */
    public static <T> T copy(Object o, Class<T> cls, String ...ignoreProperty){
        T result = getInstance(cls);
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

    public static <T> List<T> copyList(List<Object> objects, Class<T> cls, String ...ignoreProperty){
        List<T> resultList = new ArrayList<>(objects.size());
        for (Object o : objects){
            T t = copy(o,cls,ignoreProperty);
            resultList.add(t);
        }
        return resultList;
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

    /**
     *
     * @param object 检查目标 check target
     * @param ignoreProperty 忽略字段 ignore field
     * @return 没有null字段返回true ，有null字段返回false
     */
    public static boolean isNullCheck(Object object,String ...ignoreProperty){
        Class<?> cls = object.getClass();
        Set<String> fieldNames = Arrays.stream(cls.getDeclaredFields()).map(Field::getName).collect(Collectors.toSet());
        if(null != ignoreProperty && ignoreProperty.length>0){
            Arrays.asList(ignoreProperty).forEach(fieldNames::remove);
        }
        for (String name : fieldNames){
            Field field;
            try {
               field = cls.getDeclaredField(name);
            } catch (NoSuchFieldException e) {
                throw new RuntimeException(e.getMessage());
            }
            field.setAccessible(true);
            try {
                if(null == field.get(object)){
                    return false;
                }
            } catch (IllegalAccessException e) {
                throw new ConvertException("can not access property:'"+name+"', "+e.getMessage());
            }
        }
        return true;
    }

    public static boolean isBaseType(Object o){
        Class<?> cls = o.getClass();
        return cls.isInstance(Integer.class) || cls.isInstance(Byte.class) || cls.isInstance(Long.class)
                || cls.isInstance(String.class) || cls.isInstance(Float.class) || cls.isInstance(Double.class) || cls.isInstance(Boolean.class);
    }

    public static String toJSONString(Object o,String ...ignoreProperty){
        String json = "";
        try {
            json = json.concat(beanToJSON(o,ignoreProperty));
        } catch (IllegalAccessException e) {
            throw new ConvertException("can not access property in '"+o.getClass().getName()+"', "+e.getMessage());
        }
        json =json.concat("");
        return json;
    }

    private static String beanToJSON(Object o, String ...ignoreProperty) throws IllegalAccessException {
        String body = "{";
        Class<?> cls = o.getClass();
        List<String> fieldNames = Arrays.stream(cls.getDeclaredFields()).map(Field::getName).collect(Collectors.toList());
        if(null != ignoreProperty && ignoreProperty.length>0){
            Arrays.asList(ignoreProperty).forEach(fieldNames::remove);
        }
        for (String fieldName : fieldNames){
            Field field;
            try {
                field = cls.getDeclaredField(fieldName);
            } catch (NoSuchFieldException e) {
                throw new ConvertException("can not convert, no property name :'"+fieldName+"'found in '"+cls.getName()+"'");
            }
            body = body.concat("\"").concat(field.getName()).concat("\":");
            field.setAccessible(true);
            body = body.concat(fieldToJSON(field,o));
        }
        body = body.substring(0, body.length() - 1);
        body = body.concat("}");
        return body;
    }

    private static String fieldToJSON(Field field,Object o) throws IllegalAccessException {
        String body = "";
        if(field.get(o) instanceof Collection){
            body = body.concat(collectionToJSON(field,o));
        }else if(field.get(o) instanceof Map){
            body = body.concat(mapToJSON(field,o));
        }else if (field.get(o) instanceof Arrays){
            body = body.concat(arrayToJSON(field,o));
        }else if(isBaseType(field.getType())){
                body = body.concat(baseToJSON(field,o));
        }else {
            Object object = field.get(o);
            body = body.concat(beanToJSON(object));
        }
        return body;
    }

    private static String collectionToJSON(Field field, Object o) throws IllegalAccessException {
        String body = "[";
        Collection<?> collection = (Collection<?>) field.get(o);
        List<?> dataList = collection.stream().toList();
        if(null == dataList || dataList.size()==0){
            body = body.concat("],");
            return body;
        }
        Class<?> cls = dataList.get(0).getClass();
        Iterator<?> iterator = dataList.listIterator();
        if(cls.isInstance("")) {
            while (iterator.hasNext()) {
                body = body.concat("\"").concat(iterator.next().toString()).concat("\",");
            }
        }else {
            while (iterator.hasNext()) {
                body = body.concat(iterator.next().toString()).concat(",");
            }
        }
        body = body.substring(0,body.length()-1);
        body = body.concat("],");
        return body;
    }

    private static String arrayToJSON(Field field, Object o) throws IllegalAccessException {
        String body = "[";
        Object[] array = (Object[]) field.get(o);
        for (Object value : array) {
            if (value instanceof String) {
                body = body.concat("\"").concat(value.toString()).concat("\",");
            } else {
                body = body.concat(value.toString()).concat(",");
            }
        }
        body = body.substring(0,body.length()-1);
        body = body.concat("],");
        return body;
    }

    private static String mapToJSON(Field field, Object o) throws IllegalAccessException {
        String body = "{";
        Map<?,?> map = (Map<?, ?>) field.get(o);
        for (Object key : map.keySet()){
            body = body.concat("\"").concat(key.toString()).concat("\":");
            if(map.get(key) instanceof Integer){
                body = body.concat(((Integer) map.get(key)).toString()).concat(",");
            }else {
                body = body.concat("\"").concat(map.get(key).toString()).concat("\",");
            }
        }
        body = body.substring(0,body.length()-1);
        body = body.concat("},");
        return body;
    }

    private static String baseToJSON(Field field, Object o) throws IllegalAccessException {
        String body = "";
        Class<?> cls = field.getType();
        Object value = field.get(o);
        if(null == value){
            body = body.concat("null,");
            return body;
        }
        if(cls.isInstance("")){
            body = body.concat("\"").concat(field.get(o).toString()).concat("\",");
        }else {
            body = body.concat(field.get(o).toString()).concat(",");
        }
        return body;
    }



}
