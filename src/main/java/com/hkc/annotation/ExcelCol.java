package com.hkc.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author hukunchi
 * @Title: excel导出，列属性
 * 注：字段类型要用封装类
 * @Package ${package_name}
 * @Description: ${todo}
 * @date 2018/11/815:27
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelCol {
    //排序，显示地几列
    public String sort();
    //列名
    public String name();
    /***
     *格式化:
     *1.时间格式：  yyyy-MM-dd
     *2.数字类型保留小数位数：accuracy:2
     *3.百分比,会在后面追加%,保留几位小数:percent:2
     *4.百分比,直接在后面追加%:percent:-1
     *目前只支持这几种
     */
    public String format() default "";
}
