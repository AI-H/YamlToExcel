package com.feikebuqu.filedemo.yamlfile;

import com.feikebuqu.filedemo.bean.AlterBean;
import com.feikebuqu.filedemo.excelutils.ExcelUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.yaml.snakeyaml.Yaml;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class YamlDemo {
    public static void main(String[] args) throws Exception {
        Yaml yaml = new Yaml();
        Object load = null;
        Map peometheus =null;
        ArrayList<AlterBean> lists = new ArrayList<>();
        ArrayList<AlterBean> alters = new ArrayList<>();
        try {
            load = yaml.load(new FileInputStream("D:\\ideaWorkSpace\\filedemo\\src\\main\\resources\\prometheus-rules.yaml"));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        if (load == null) {
            System.out.println("文件是空");
        }else{
            peometheus = (Map)load;
//            System.out.println(peometheus.get("apiVersion"));
            Map spec = (Map)peometheus.get("spec");
            List groups = (List)spec.get("groups");

            for (Object group : groups) {
                Map name = (Map)group;
                String nameValue = (String)name.get("name");
                List rules = (List)name.get("rules");
                for (Object rule : rules) {
                    Map ruleValue = (Map)rule;
                    Object alert = ruleValue.get("alert");
                    if(alert == null){
                        //不包含alter，前几个expr
                        String expr = (String)ruleValue.get("expr");
                        Object labelsObj = ruleValue.get("labels");
                        String labels = "";
                        if (labelsObj != null){
                            labels = labelsObj.toString();
                        }
                        String record = (String)ruleValue.get("record");

                        AlterBean alterBean = new AlterBean();
                        alterBean.setName(nameValue);
                        alterBean.setExpr(expr);
                        alterBean.setRecord(record);
                        alterBean.setLabels(labels);
                        lists.add(alterBean);


                    }else{
                        //包含alter
                        String alterValue = (String) alert;
                        Map annotations = (Map)ruleValue.get("annotations");
                        String message = (String)annotations.get("message");
                        String runbook_url = (String)annotations.get("runbook_url");
                        String expr = (String)ruleValue.get("expr");
                        String forV = (String)ruleValue.get("for");
                        Object labelsObj = ruleValue.get("labels");
                        String labels = "";
                        if (labelsObj != null){
                            labels = labelsObj.toString();
                        }

                        AlterBean alterBean = new AlterBean();
                        alterBean.setName(nameValue);
                        alterBean.setExpr(expr);
                        alterBean.setAlterValue(alterValue);
                        alterBean.setMessage(message);
                        alterBean.setRunbook_url(runbook_url);
                        alterBean.setForV(forV);
                        alterBean.setLabels(labels);
                        alters.add(alterBean);
                    }
                }
            }

        }
    //将数据输出到excel表格
        writeExcel(alters,"D:\\alters.xls",AlterBean.class);
        writeExcel(lists,"D:\\lists.xls",AlterBean.class);

    }

    public static<E>  void writeExcel(List<E> list,String filename,Class<E> eclass)throws  Exception{
        Workbook wb = ExcelUtil.createTsWorkBook(list,eclass);
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        wb.write(os);
        byte[] content = os.toByteArray();
        InputStream is = new ByteArrayInputStream(content);


        FileOutputStream outputStream = new FileOutputStream(filename);
        byte[] bytes = new byte[2 * 1024];
        int read;
        while (-1 != (read=is.read(bytes))){
            outputStream.write(bytes,0,read);
        }
        outputStream.close();
        is.close();
    }
}
