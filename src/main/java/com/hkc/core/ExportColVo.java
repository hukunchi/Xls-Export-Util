package com.hkc.core;

import java.io.Serializable;

/**
 * @author hukunchi
 * @Title: ${file_name}
 * @Package ${package_name}
 * @Description: ${todo}
 * @date 2018/11/815:28
 */
public class ExportColVo implements Serializable {

    private static final long serialVersionUID= 1L;

    // 序号
    private Integer sort;
    // 列名称
    private String colName;
    // 字段名
    private String fieldName;
    /***
     * 格式化:
     * 1.时间格式： date:yyyy-MM-dd
     * 2.数字类型保留小数位数：accuracy:2
     * 3.百分比,保留几位小数:percent:2
     * 4.百分比,追加百分号：percent:-1
     * 5.日期类型：date:yyyy-MM-dd
     *
     * 目前只支持这几种
     */
    private String format;


    public ExportColVo(Integer sort, String colName, String fieldName, String format) {
        super();
        this.sort = sort;
        this.colName = colName;
        this.fieldName = fieldName;
        this.format = format;
    }

    public Integer getSort() {
        return sort;
    }

    public void setSort(Integer sort) {
        this.sort = sort;
    }

    public String getColName() {
        return colName;
    }

    public void setColName(String colName) {
        this.colName = colName;
    }

    public String getFieldName() {
        return fieldName;
    }

    public void setFieldName(String fieldName) {
        this.fieldName = fieldName;
    }

    public String getFormat() {
        return format;
    }

    public void setFormat(String format) {
        this.format = format;
    }

    @Override
    public String toString() {
        return "ExportColVo [sort=" + sort + ", colName=" + colName + ", fieldName=" + fieldName + ", format=" + format
                + "]";
    }


}
