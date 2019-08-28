package com.feikebuqu.filedemo.bean;

import java.util.List;

/**
 * Created by HuangLiTe on 2019/5/16.
 */
public class ExcelBean<E> {
    private List<E> list;
    private Class<E> clazz;
    private String tableName;
    private String lastRowName;
    private String lastRowValue;

    public String getLastRowName() {
        return lastRowName;
    }

    public void setLastRowName(String lastRowName) {
        this.lastRowName = lastRowName;
    }

    public String getLastRowValue() {
        return lastRowValue;
    }

    public void setLastRowValue(String lastRowValue) {
        this.lastRowValue = lastRowValue;
    }

    public List<E> getList() {
        return list;
    }

    public void setList(List<E> list) {
        this.list = list;
    }

    public Class<E> getClazz() {
        return clazz;
    }

    public void setClazz(Class<E> clazz) {
        this.clazz = clazz;
    }

    public String getTableName() {
        return tableName;
    }

    public void setTableName(String tableName) {
        this.tableName = tableName;
    }
}
