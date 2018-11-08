package com.hkc.util;

import com.hkc.annotation.ExcelCol;

import java.math.BigDecimal;

/**
 * @author hukunchiwj53006@touna.cn
 * @Title: ${file_name}
 * @Package ${package_name}
 * @Description: ${todo}
 * @date 2018/11/816:37
 */
public class DemoVo {
    @ExcelCol(name = "姓名",sort = "1")
    private String name;
    @ExcelCol(name = "年龄",sort = "2")
    private String age;
    @ExcelCol(name = "地址",sort = "3")
    private String addr;
    @ExcelCol(name = "余额",sort = "4")
    private BigDecimal amout;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getAge() {
        return age;
    }

    public void setAge(String age) {
        this.age = age;
    }

    public String getAddr() {
        return addr;
    }

    public void setAddr(String addr) {
        this.addr = addr;
    }

    public BigDecimal getAmout() {
        return amout;
    }

    public void setAmout(BigDecimal amout) {
        this.amout = amout;
    }
}
