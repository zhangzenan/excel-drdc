package com.imooc;

import org.apache.commons.lang3.StringUtils;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.jdom.Attribute;
import org.jdom.Document;
import org.jdom.Element;
import org.jdom.JDOMException;
import org.jdom.input.SAXBuilder;

import java.io.File;
import java.io.IOException;
import java.util.List;

public class CreateTemplate {
    public static void main(String[] args) {
        //获取解析xml文件路径
        String path = System.getProperty("user.dir") + "/src/main/java/com/imooc/template/student.xml";
        File file = new File(path);
        SAXBuilder builder = new SAXBuilder();
        try {
            //解析xml文件
            Document parse = builder.build(file);
            //创建Excel
            HSSFWorkbook wb = new HSSFWorkbook();
            //创建sheet
            HSSFSheet sheet = wb.createSheet("Sheet0");
            //获取xml文件根节点
            Element root = parse.getRootElement();
            //获取模板名称
            String templateName = root.getAttribute("name").getValue();

            int rownum = 0;
            int column = 0;
            //设置列宽
            Element colgroup = root.getChild("colgroup");
            setColumnWidth(sheet, colgroup);

            //设置标题
            Element title = root.getChild("title");
            List<Element> trs = title.getChildren("tr");
            for (int i = 0; i < trs.size(); i++) {
                Element tr = trs.get(i);
                List<Element> tds = tr.getChildren("td");
                HSSFRow row = sheet.createRow(rownum);
                HSSFCellStyle cellStyle = wb.createCellStyle();
                //cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);

                for (column = 0; column < tds.size(); column++) {
                    Element td = tds.get(column);
                    HSSFCell cell = row.createCell(column);
                    Attribute rowSpan = td.getAttribute("rowspan");
                    Attribute colSpan = td.getAttribute("colspan");
                    Attribute value = td.getAttribute("value");
                    if (value != null) {
                        String val = value.getValue();
                        cell.setCellValue(val);
                        int rspan = rowSpan.getIntValue() - 1;
                        int cspan = colSpan.getIntValue() - 1;

                        //设置字体
                        HSSFFont font = wb.createFont();
                        font.setFontName("仿宋_GB2312");
                        //font.setBoldweight(HSSFFont.BLODWEIGHT_BOLD);//字体加粗
                        font.setFontHeight((short) 12);
                        cellStyle.setFont(font);
                        cell.setCellStyle(cellStyle);

                        //合并单元格居中
                        sheet.addMergedRegion(new CellRangeAddress(rspan, rspan, 0, cspan));
                    }


                }

            }


        } catch (JDOMException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println(path);

    }

    private static void setColumnWidth(HSSFSheet sheet, Element colgroup) {
        List<Element> cols = colgroup.getChildren("col");
        for (int i = 0; i < cols.size(); i++) {
            Element col = cols.get(i);
            Attribute width = col.getAttribute("width");
            String unit = width.getValue().replaceAll("[0-9,\\.]", "");
            String value = width.getValue().replace(unit, "");
            int v = 0;
            if (StringUtils.isBlank(unit) || "px".endsWith(unit)) {
                v = Math.round(Float.parseFloat(value) * 37F);
            } else if ("em".endsWith(unit)) {
                v = Math.round(Float.parseFloat(value) * 267.5F);
            }
            sheet.setColumnWidth(i, v);
        }
    }
}
