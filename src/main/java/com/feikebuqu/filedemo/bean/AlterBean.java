package com.feikebuqu.filedemo.bean;

import java.io.Serializable;

public class AlterBean implements Serializable{


    private static final long serialVersionUID = -1795237607533763535L;
    @ExcelField("名称")
    private String name;
    @ExcelField("指标表达式")
    private String expr;
    @ExcelField("record记录")
    private String record;
    @ExcelField("预警名称")
    private String alterValue;
    @ExcelField("message预警信息")
    private String message;
    @ExcelField("参考链接")
    private String runbook_url;
    @ExcelField("时间段")
    private String forV;
    @ExcelField("标签")
    private String labels;


    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getExpr() {
        return expr;
    }

    public void setExpr(String expr) {
        this.expr = expr;
    }

    public String getRecord() {
        return record;
    }

    public void setRecord(String record) {
        this.record = record;
    }

    public String getAlterValue() {
        return alterValue;
    }

    public void setAlterValue(String alterValue) {
        this.alterValue = alterValue;
    }

    public String getMessage() {
        return message;
    }

    public void setMessage(String message) {
        this.message = message;
    }

    public String getRunbook_url() {
        return runbook_url;
    }

    public void setRunbook_url(String runbook_url) {
        this.runbook_url = runbook_url;
    }

    public String getForV() {
        return forV;
    }

    public void setForV(String forV) {
        this.forV = forV;
    }

    public String getLabels() {
        return labels;
    }

    public void setLabels(String labels) {
        this.labels = labels;
    }
}
