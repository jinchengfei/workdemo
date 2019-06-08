/**
 * Copyright (C), 2015-2019, XXX有限公司
 * FileName: ExcelChartUtil2
 * Author:   jcf
 * Date:     2019/5/31 23:00
 * Description:
 * History:
 * <author>          <time>          <version>          <desc>
 * 作者姓名           修改时间           版本号              描述
 */

package com.example.exportexcel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelChartUtil2 {
    private static SXSSFWorkbook wb = new SXSSFWorkbook();

    public static void main(String[] args) {
        // 字段名
        List<String> fldNameArr = new ArrayList<String>();
        // 标题
        List<String> titleArr = new ArrayList<String>();
        titleArr.add("指标名称");
        titleArr.add("数值");

        List<String> xStrs = new ArrayList<>();
        List<String> yStrs = new ArrayList<>();
        int year0 = 2012;
        int month = 1;
        for (int i = 0; i <6; i++) {
            if(month>12){
                month = 1;
                year0 ++;
            }
            String ym = String.valueOf(year0);
            if(month<10){
                ym = ym + "0"+String.valueOf(month);
            }else {
                ym = ym+String.valueOf(month);
            }
            month++;
            xStrs.add(ym);
            if(i%2 == 0 ){
                yStrs.add(-Math.floor(Math.random() * 100) + "%");
            }else{
                yStrs.add(Math.floor(Math.random() * 100) + "%");
            }
        }

        List<Map<String, Object>> dataList = new ArrayList<Map<String, Object>>();
        for (int j = 0; j<xStrs.size();j++) {
            Map<String, Object> hashMap = new HashMap<>();
            hashMap.put("value1", xStrs.get(j));

            hashMap.put("value2",yStrs.get(j));
            dataList.add(hashMap);
        }
        fldNameArr.add("value1");
        fldNameArr.add("value2");
        ExcelChartUtil2 ecu = new ExcelChartUtil2();
        try {
            // 创建折线图
            ecu.createTimeXYChar(titleArr, fldNameArr, dataList);
            //导出到文件
            FileOutputStream out = new FileOutputStream(new File("D:/CXF/" + System.currentTimeMillis() + ".xls"));
            wb.write(out);
            out.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    /**
     * 创建折线图
     *
     * @throws
     */
    public void createTimeXYChar(List<String> titleArr, List<String> fldNameArr, List<Map<String, Object>> dataList) {
        List<SXSSFSheet> sheets = new ArrayList<>();
        // 创建一个sheet页
        SXSSFSheet sheet0 = wb.createSheet("数据折现图");
        SXSSFSheet sheet1 = wb.createSheet("数据源");
        sheets.add(sheet0);
        sheets.add(sheet1);

        // 第二个参数折线图类型:line=普通折线图,line-bar=折线+柱状图
        boolean result = drawSheet3Map(sheets, "line", fldNameArr, dataList, titleArr);
        System.out.println("生成折线图(折线图or折线图-柱状图)-->" + result);

    }

    /**
     * 生成折线图
     *
     * @param sheets
     *            页签
     * @param type
     *            类型
     * @param fldNameArr
     *            X轴标题
     * @param dataList
     *            填充数据
     * @param titleArr
     *            图例标题
     * @return
     */
    private boolean drawSheet3Map(List<SXSSFSheet> sheets, String type, List<String> fldNameArr,
                                  List<Map<String, Object>> dataList, List<String> titleArr) {
        boolean result = false;
        SXSSFSheet  sheet1 = sheets.get(1);
        // 获取sheet名称
        String sheetName = sheet1.getSheetName();
        //创建表格
        result = drawSheet0Table(sheet1, titleArr, fldNameArr, dataList);

        SXSSFSheet  sheet0 = sheets.get(0);
        // 创建一个画布
        Drawing<?> drawing = sheet0.createDrawingPatriarch();
        // 画一个图区域
        // 前四个默认0，从第0行到第30行,从第0列到第22列的区域   jcf
        ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 0, 26, 40);
        // 创建一个chart对象
        Chart chart = drawing.createChart(anchor);
        CTChart ctChart = ((XSSFChart) chart).getCTChart();
        CTPlotArea ctPlotArea = ctChart.getPlotArea();
        // 折线图
        CTLineChart ctLineChart = ctPlotArea.addNewLineChart();
        CTBoolean ctBoolean = ctLineChart.addNewVaryColors();
        ctLineChart.addNewGrouping().setVal(STGrouping.STANDARD);
        // 创建序列,并且设置选中区域
        for (int i = 0; i < fldNameArr.size() - 1; i++) {
            CTLineSer ctLineSer = ctLineChart.addNewSer();
            CTSerTx ctSerTx = ctLineSer.addNewTx();
            // 图例区
            CTStrRef ctStrRef = ctSerTx.addNewStrRef();
            // 选定区域第0行,第1,2,3列标题作为图例 //1 2 3
            //String legendDataRange = new CellRangeAddress(0, 0, i + 1, i + 1).formatAsString(sheetName, true);
            String legendDataRange = new CellRangeAddress(0, 0, i, i ).formatAsString(sheetName, true);
            ctStrRef.setF(legendDataRange);
            ctLineSer.addNewIdx().setVal(i);
            // 横坐标区
            CTAxDataSource cttAxDataSource = ctLineSer.addNewCat();
            ctStrRef = cttAxDataSource.addNewStrRef();
            // 选第0列,第1-dataList.size()行作为横坐标区域
            String axisDataRange = new CellRangeAddress(1, dataList.size(), 0, 0).formatAsString(sheetName, true);
            ctStrRef.setF(axisDataRange);
            // 数据区域
            CTNumDataSource ctNumDataSource = ctLineSer.addNewVal();
            CTNumRef ctNumRef = ctNumDataSource.addNewNumRef();
            // 选第1- dataList.size()行,第1列作为数据区域
            String numDataRange = new CellRangeAddress(1, dataList.size(), i + 1, i + 1).formatAsString(sheetName,
                    true);
            System.out.println(numDataRange);
            ctNumRef.setF(numDataRange);
            // 设置标签格式
            ctBoolean.setVal(false);
            CTDLbls newDLbls = ctLineSer.addNewDLbls();
            newDLbls.setShowLegendKey(ctBoolean);
            ctBoolean.setVal(true);
            newDLbls.setShowVal(ctBoolean);
            ctBoolean.setVal(false);
            newDLbls.setShowCatName(ctBoolean);
            newDLbls.setShowSerName(ctBoolean);
            newDLbls.setShowPercent(ctBoolean);
            newDLbls.setShowBubbleSize(ctBoolean);
            newDLbls.setShowLeaderLines(ctBoolean);
            // 是否是平滑曲线
            CTBoolean addNewSmooth = ctLineSer.addNewSmooth();
            addNewSmooth.setVal(false);
            // 是否是堆积曲线
            CTMarker addNewMarker = ctLineSer.addNewMarker();
            CTMarkerStyle addNewSymbol = addNewMarker.addNewSymbol();
            addNewSymbol.setVal(STMarkerStyle.NONE);
        }
        // telling the BarChart that it has axes and giving them Ids
        ctLineChart.addNewAxId().setVal(123456);
        ctLineChart.addNewAxId().setVal(123457);
        // cat axis
        CTCatAx ctCatAx = ctPlotArea.addNewCatAx();
        ctCatAx.addNewAxId().setVal(123456); // id of the cat axis
        CTScaling ctScaling = ctCatAx.addNewScaling();
        ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
        ctCatAx.addNewAxPos().setVal(STAxPos.B);
        ctCatAx.addNewCrossAx().setVal(123457); // id of the val axis
        ctCatAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);
        // val axis
        CTValAx ctValAx = ctPlotArea.addNewValAx();
        ctValAx.addNewAxId().setVal(123457); // id of the val axis
        ctScaling = ctValAx.addNewScaling();
        ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
        ctValAx.addNewAxPos().setVal(STAxPos.L);
        ctValAx.addNewCrossAx().setVal(123456); // id of the cat axis
        ctValAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);
        // 是否删除主左边轴
        ctValAx.addNewDelete().setVal(false);
        // 是否删除横坐标
        ctCatAx.addNewDelete().setVal(false);
        CTLegend ctLegend = ctChart.addNewLegend();
        ctLegend.addNewLegendPos().setVal(STLegendPos.B);
        ctLegend.addNewOverlay().setVal(false);
        return result;
    }
    /**
     * 生成数据表
     *
     * @param sheet
     *            sheet页对象
     * @param titleArr
     *            表头字段
     * @param fldNameArr
     *            左边标题字段
     * @param dataList
     *            数据
     * @return 是否生成成功
     */
    private boolean drawSheet0Table(SXSSFSheet sheet, List<String> titleArr, List<String> fldNameArr,
                                    List<Map<String, Object>> dataList) {
        // 测试时返回值
        boolean result = true;
        // 初始化表格样式
        List<CellStyle> styleList = tableStyle();
        // 根据数据创建excel第一行标题行
        SXSSFRow row0 = sheet.createRow(0);
        for (int i = 0; i < titleArr.size(); i++) {
            // 设置标题
            row0.createCell(i).setCellValue(titleArr.get(i));
            // 设置标题行样式
            row0.getCell(i).setCellStyle(styleList.get(0));
        }
        // 填充数据
        for (int i = 0; i < dataList.size(); i++) {
            // 获取每一项的数据
            Map<String, Object> data = dataList.get(i);
            // 设置每一行的字段标题和数据
            SXSSFRow rowi = sheet.createRow(i + 1);
            for (int j = 0; j < data.size(); j++) {
                // 判断是否是标题字段列
                if (j == 0) {
                    rowi.createCell(j).setCellValue((String) data.get("value" + (j + 1)));
                    // 设置左边字段样式
                    sheet.getRow(i + 1).getCell(j).setCellStyle(styleList.get(0));
                } else {
                    //rowi.createCell(j).setCellValue(Double.valueOf((String) data.get("value" + (j + 1))));
                    rowi.createCell(j).setCellValue((String) data.get("value" + (j + 1)));


                    // 设置数据样式
                    sheet.getRow(i + 1).getCell(j).setCellStyle(styleList.get(2));
                }
            }
        }
        return result;
    }

    /**
     * 生成表格样式
     *
     * @return
     */
    private static List<CellStyle> tableStyle() {
        List<CellStyle> cellStyleList = new ArrayList<CellStyle>();
        // 样式准备
        // 标题样式
        CellStyle style = wb.createCellStyle();
        style.setFillForegroundColor(IndexedColors.ROYAL_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN); // 下边框
        style.setBorderLeft(BorderStyle.THIN);// 左边框
        style.setBorderTop(BorderStyle.THIN);// 上边框
        style.setBorderRight(BorderStyle.THIN);// 右边框
        style.setAlignment(HorizontalAlignment.CENTER);
        cellStyleList.add(style);

        CellStyle style1 = wb.createCellStyle();
        style1.setBorderBottom(BorderStyle.THIN); // 下边框
        style1.setBorderLeft(BorderStyle.THIN);// 左边框
        style1.setBorderTop(BorderStyle.THIN);// 上边框
        style1.setBorderRight(BorderStyle.THIN);// 右边框
        style1.setAlignment(HorizontalAlignment.CENTER);
        cellStyleList.add(style1);

        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setBorderTop(BorderStyle.THIN);// 上边框
        cellStyle.setBorderBottom(BorderStyle.THIN); // 下边框
        cellStyle.setBorderLeft(BorderStyle.THIN);// 左边框
        cellStyle.setBorderRight(BorderStyle.THIN);// 右边框
        cellStyle.setAlignment(HorizontalAlignment.CENTER);// 水平对齐方式
        // cellStyle.setVerticalAlignment(VerticalAlignment.TOP);//垂直对齐方式
        cellStyleList.add(cellStyle);
        return cellStyleList;
    }
}
