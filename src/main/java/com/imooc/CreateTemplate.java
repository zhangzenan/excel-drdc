package com.imooc;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.jdom.Document;
import org.jdom.Element;
import org.jdom.JDOMException;
import org.jdom.input.SAXBuilder;

import java.io.File;
import java.io.IOException;

public class CreateTemplate {
    public static void main(String[] args) {
        //获取解析xml文件路径
        String path = System.getProperty("user.dir") + "/src/main/java/com/imooc/template/student.xml";
        File file = new File(path);
        SAXBuilder builder = new SAXBuilder();
        try {
            //解析xml文件
            Document parse= builder.build(file);
            //创建Excel
            HSSFWorkbook wb=new HSSFWorkbook();
            //创建sheet
            HSSFSheet sheet=wb.createSheet("Sheet0");
            //获取xml文件根节点
            Element root=parse.getRootElement();
            //获取模板名称
            String templateName=root.getAttribute("name").getValue();

            int rownum=0;
            int column=0;
            //设置列宽


        } catch (JDOMException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println(path);

    }
}
