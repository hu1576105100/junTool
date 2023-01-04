package com.jun.tool;


import java.beans.BeanInfo;
import java.beans.IntrospectionException;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.time.LocalDateTime;
import java.util.*;

public class BeanUtil {


    /**
     * 判空
     */
    public static boolean isEmpty(Object obj){
        if(obj==null){
            return true;
        }
        if (obj instanceof Map) {
            return ((Map<?, ?>) obj).isEmpty();
        } else if (obj instanceof List) {
            return ((List< ? >) obj).isEmpty();
        } else if (obj instanceof String) {
            return ((String) obj).length() == 0;
        }
        return false;
    }
    public static boolean isNotEmpty(Object obj){
        return !isEmpty(obj);
    }


    /***
     * 将实体类转换成map
     */
    public static Map<String, Object> toMap(Object obj)  {
        Class type = obj.getClass();
        Map<String, Object> returnMap = new HashMap<>();
        BeanInfo beanInfo = null;
        try {
            beanInfo = Introspector.getBeanInfo(type);
            PropertyDescriptor[] propertyDescriptors = beanInfo.getPropertyDescriptors();
            for (int i = 0; i < propertyDescriptors.length; i++) {
                PropertyDescriptor descriptor = propertyDescriptors[i];
                String propertyName = descriptor.getName();

                if (!propertyName.equals("class")) {
                    Method readMethod = descriptor.getReadMethod();
                    Object result = readMethod.invoke(obj);
                    if (result != null) {
                        if(descriptor.getPropertyType().equals(LocalDateTime.class)){
                            returnMap.put(propertyName, DateUtils.toString((LocalDateTime)result,DateUtils.date));
                        }else{
                            returnMap.put(propertyName, result);
                        }
                    } else {
                        returnMap.put(propertyName, "");
                    }
                }
            }
        } catch (IntrospectionException | InvocationTargetException | IllegalAccessException e) {
            e.printStackTrace();
        }
        return returnMap;
    }

}
