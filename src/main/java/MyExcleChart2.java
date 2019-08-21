import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeProperties;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @program: poi
 * @description:
 * @author: sickle
 * @create: 2019-08-19 16:13
 **/
public class MyExcleChart2 {

    public static void doWork(List<String> title, List<String> styleList, Map<String, List<Object>> day2ColValueList, File file,
                              String sheetName, XSSFWorkbook wb, int dateNum) throws IOException {
        OutputStream out = null;
        try {
            int sheetIndex = wb.getSheetIndex(sheetName);
            if (sheetIndex >= 0) {
                wb.removeSheetAt(sheetIndex);
            }
            int sheetNum = wb.getNumberOfSheets();
            XSSFSheet sheet = wb.createSheet();
            wb.setSheetName(sheetNum, sheetName);
            out = new FileOutputStream(file);

            //设置内容样式
            XSSFCellStyle style = setBorder(wb);

            //设置表头字体
            XSSFFont font = wb.createFont();
            font.setBold(true);    //加粗
            //设置表头样式
            XSSFCellStyle headStyle = setBorder(wb);
            headStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headStyle.setFont(font);

            //隐藏列
            hiddenColumn(sheet, sheetName);

            Row row;
            Cell cell = null;
            row = sheet.createRow(0);
            //写入表头
            int titleColIndex = 0;
            for (String t : title) {
                cell = row.createCell((short) titleColIndex);
                cell.setCellValue(t);
                cell.setCellStyle(headStyle);
                titleColIndex++;
            }
            //写入数据
            int rowIndex = 1;
            for (String key : day2ColValueList.keySet()) {
                row = sheet.createRow(rowIndex);

                List<Object> dataList = day2ColValueList.get(key);
                cell = row.createCell(0);
                cell.setCellStyle(style);
                cell.setCellValue(rowIndex);

                int cellIndex = 1;
                for (Object s : dataList) {
                    //填充单元格
                    String cellstyle = styleList.get(dataList.indexOf(s) + 1);
                    cell = row.createCell(cellIndex);
                    cell.setCellStyle(style);
                    //此处可以对特殊的行进行处理
                    if ("speciaRowName".equals(key) && cellIndex > 9) {
                        cell = row.createCell(0);
                        cell.setCellStyle(style);
                        cell.setCellValue("speciaRowName");
                        cell = row.createCell(1);
                        cell.setCellStyle(style);
                        cell.setCellValue("");
                        cell = row.createCell(2);
                        cell.setCellStyle(style);
                        cell.setCellValue("");
                        cell = row.createCell(3);
                        cell.setCellStyle(style);
                        cell.setCellValue("");
                        cell = row.createCell(4);
                        cell.setCellStyle(style);
                        cell.setCellValue("");
                        cell = row.createCell(cellIndex);
                        cell.setCellStyle(style);
                        double dble = (double) s;
                        cell.setCellValue(dble);
                    } else if ("int".equals(cellstyle) && null != s) {
                        int num = (int) s;
                        cell.setCellValue(num);
                    } else if ("double".equals(cellstyle) && null != s) {
                        double dble = (double) s;
                        cell.setCellValue(dble);
                    } else {
                        cell.setCellValue(null == s ? "" : (String) s);
                    }
                    cellIndex++;
                }
                rowIndex++;
            }
            //绘制图表
            if (day2ColValueList.size() > 0) {
                drawChart(sheet, sheetName, day2ColValueList, titleColIndex, dateNum);
            }
            wb.write(out);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                out.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private static void drawChart(XSSFSheet sheet, String sheetName, Map<String, List<Object>> day2ColValueList,
                                  int titleSize, int dateNum) {
        Map<String, Integer> paramMap = new HashMap<String, Integer>();// 折线图入参
        Map<String, Object> pieParamMap = new HashMap<String, Object>();// 扇形图入参
        //此处可以分sheet处理数据
        if ("sheetname1".equals(sheetName)) {//sheetname1，需要显示折线图和扇形图
            paramMap.put("numstartcol", 5);
            paramMap.put("numendcol", 5 + dateNum - 1);
            paramMap.put("prostartcol", 5 + dateNum);
            paramMap.put("proendcol", 5 + dateNum + dateNum - 1);
            // 业务汇总表，需要统计身份核查、高清人像、银行卡、手机实名四个业务的折线图
            for (String key : day2ColValueList.keySet()) {
                if ("类型1业务1".equals(key)) {// 人像比对认证走势折线图
                    // 折线图x轴单位起止列（numstartcol，numendcol），以及计费笔数数据所在行列
                    paramMap.put("col1", 0);
                    paramMap.put("col2", 7);
                    paramMap.put("row1", day2ColValueList.size() + 2);
                    paramMap.put("row2", day2ColValueList.size() + 19);
                    paramMap.put("numstartrow", 1);
                    paramMap.put("numendrow", 1);
                    paramMap.put("numstartrow2", 2);
                    paramMap.put("numendrow2", 2);
                    // 折线图净收入所在行列
                    paramMap.put("prostartrow", 1);
                    paramMap.put("proendrow", 1);
                    paramMap.put("prostartrow2", 2);
                    paramMap.put("proendrow2", 2);
                    drawLineChart(sheet, "业务1", paramMap);
                } else if ("类型1业务2".equals(key)) {// 银行卡认证走势折线图
                    paramMap.put("col1", 8);
                    paramMap.put("col2", 15);
                    paramMap.put("row1", day2ColValueList.size() + 2);
                    paramMap.put("row2", day2ColValueList.size() + 19);
                    paramMap.put("numstartrow", 3);
                    paramMap.put("numendrow", 3);
                    paramMap.put("numstartrow2", 4);
                    paramMap.put("numendrow2", 4);
                    paramMap.put("prostartrow", 3);
                    paramMap.put("proendrow", 3);
                    paramMap.put("prostartrow2", 4);
                    paramMap.put("proendrow2", 4);
                    drawLineChart(sheet, "业务2", paramMap);
                }
            }
            //柱状图
            paramMap.put("col1", 15);
            paramMap.put("col2", 23);
            paramMap.put("row1", day2ColValueList.size() + 2);
            paramMap.put("row2", day2ColValueList.size() + 19);
            paramMap.put("numstartrow", 1);
            paramMap.put("numendrow", 1);
            paramMap.put("prostartrow", 1);
            paramMap.put("proendrow", 1);
//			drawBarChart(sheet, "业务1", paramMap);

            // 扇形图
            pieParamMap.put("col1", 0);
            pieParamMap.put("col2", 7);
            pieParamMap.put("row1", day2ColValueList.size() + 20);
            pieParamMap.put("row2", day2ColValueList.size() + 39);
            pieParamMap.put("data1", "sheetname1!$C$1");
            pieParamMap.put("data2", "sheetname1!$C$2:$C$3");
            pieParamMap.put("data3", "sheetname1!$D$2:$D$3");
            drawPieChart(sheet, "邀请人数占比", pieParamMap);
            pieParamMap.put("col1", 8);
            pieParamMap.put("col2", 15);
            pieParamMap.put("data1", "sheetname1!$C$1");
            pieParamMap.put("data2", "sheetname1!$C$4:$C$5");
            pieParamMap.put("data3", "sheetname1!$D$4:$D$5");
            drawPieChart(sheet, "注册人数占比", pieParamMap);
            pieParamMap.put("col1", 0);
            pieParamMap.put("col2", 7);
            pieParamMap.put("row1", day2ColValueList.size() + 40);
            pieParamMap.put("row2", day2ColValueList.size() + 59);
            pieParamMap.put("data1", "sheetname1!$C$1");
            pieParamMap.put("data2", "sheetname1!$C$6:$C$7");
            pieParamMap.put("data3", "sheetname1!$D$6:$D$7");
            drawPieChart(sheet, "住宿人数占比", pieParamMap);
            pieParamMap.put("col1", 8);
            pieParamMap.put("col2", 15);
            pieParamMap.put("data1", "sheetname1!$C$1");
            pieParamMap.put("data2", "sheetname1!$C$8:$C$9");
            pieParamMap.put("data3", "sheetname1!$D$8:$D$9");
            drawPieChart(sheet, "接机人数占比", pieParamMap);
        }
    }

    /**
     * 绘制扇形图
     *
     * @param sheet    sheet
     * @param string   标题
     * @param paramMap 各种起始截止行列
     *                 col1 col2 row1 row2 图片坐标
     *                 data1  种类划分标志所在列
     *                 data2  各分类名
     *                 data3  各分类数值
     */
    private static void drawPieChart(XSSFSheet sheet, String title, Map<String, Object> pieParamMap) {
        int col1 = (int) pieParamMap.get("col1");
        int col2 = (int) pieParamMap.get("col2");
        int row1 = (int) pieParamMap.get("row1");
        int row2 = (int) pieParamMap.get("row2");
        String data1 = (String) pieParamMap.get("data1");
        String data2 = (String) pieParamMap.get("data2");
        String data3 = (String) pieParamMap.get("data3");

        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = (XSSFClientAnchor) drawing.createAnchor(0, 0, 0, 0, col1, row1, col2, row2);

        XSSFChart chart = drawing.createChart(anchor);
        chart.setTitleText(title);
        chart.setTitleOverlay(false);

        CTChart ctChart = ((XSSFChart) chart).getCTChart();
        CTPlotArea ctPlotArea = ctChart.getPlotArea();
        CTPieChart ctPieChart = ctPlotArea.addNewPieChart();
        CTBoolean ctBoolean = ctPieChart.addNewVaryColors();
        ctBoolean.setVal(true);

        CTPieSer ctPieSer = ctPieChart.addNewSer();
        CTSerTx ctSerTx = ctPieSer.addNewTx();
        CTStrRef ctStrRefTx = ctSerTx.addNewStrRef();
        ctStrRefTx.setF(data1);
        ctPieSer.addNewIdx().setVal(0);
        CTAxDataSource cttAxDataSource = ctPieSer.addNewCat();
        CTStrRef ctStrRef = cttAxDataSource.addNewStrRef();
        ctStrRef.setF(data2); // 第一行为标题
        CTNumDataSource ctNumDataSource = ctPieSer.addNewVal();
        CTNumRef ctNumRef = ctNumDataSource.addNewNumRef();
        ctNumRef.setF(data3); // 第一行为标题

        ctPieSer.addNewDLbls().addNewShowLeaderLines();// 有无此行代码,图上是否显示文字

        // legend图注
        CTLegend ctLegend = ctChart.addNewLegend();
        ctLegend.addNewLegendPos().setVal(STLegendPos.TR);
        ctLegend.addNewOverlay().setVal(true);

        ctPieSer.addNewExplosion().setVal(1);// 各块之间间隙大小
        ctPieSer.addNewOrder().setVal(0);//

        CTShapeProperties cTShapeProperties = CTShapeProperties.Factory.newInstance();
        ctPieSer.addNewSpPr().set(cTShapeProperties);
    }

    /**
     * 绘制折线图
     *
     * @param sheet            sheet
     * @param desc             横轴描述
     * @param 各种起始截止行列，包含如下内容： int numstartrow,int numendrow,int numstartcol,int numendcol,//需要绘图的计费笔数起始截止行列
     *                         int prostartrow,int proendrow,int prostartcol,int proendcol,//需要绘图的业务净收入起始截止行列
     *                         int col1,int row1,int row2 //绘图的起始行列
     */
    private static void drawLineChart(XSSFSheet sheet, String desc, Map<String, Integer> paramMap) {
        int col1 = paramMap.get("col1"), col2 = paramMap.get("col2"), row1 = paramMap.get("row1"), row2 = paramMap.get("row2"),//绘图所在坐标，默认宽度为12列
                //双折线图x轴单位起止列（numstartcol，numendcol），以及第一类数据所在行列
                numstartrow = paramMap.get("numstartrow"), numendrow = paramMap.get("numendrow"), numstartcol = paramMap.get("numstartcol"), numendcol = paramMap.get("numendcol"),
                //双折线图数据2所在行列
                prostartrow = paramMap.get("prostartrow"), proendrow = paramMap.get("proendrow"), prostartcol = paramMap.get("prostartcol"), proendcol = paramMap.get("proendcol");

        int dx1 = 0;
        int dy1 = 0;
        int dx2 = 0;
        int dy2 = 0;
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(dx1, dy1, dx2, dy2, col1, row1, col2, row2);

        XSSFChart chart = drawing.createChart(anchor);
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.TOP_RIGHT);

        // Use a category axis for the bottom axis.
        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.TOP);//底部X轴
        bottomAxis.setTitle(desc + "交易情况汇总"); // https://stackoverflow.com/questions/32010765

        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);//左侧Y轴
        leftAxis.setTitle("交易量/金额");
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

        XDDFLineChartData leftdata = (XDDFLineChartData) chart.createData(ChartTypes.LINE, bottomAxis, leftAxis);

        /////填充数据
        CellRangeAddress cellRangeAddress = new CellRangeAddress(0, 0, numstartcol, numendcol);
        XDDFDataSource<String> xs = XDDFDataSourcesFactory.fromStringCellRange(sheet, cellRangeAddress);//日期

        CellRangeAddress dataCellRangeAddress = new CellRangeAddress(numstartrow, numendrow, numstartcol, numendcol);
        XDDFNumericalDataSource<Double> ys1 = XDDFDataSourcesFactory.fromNumericCellRange(sheet, dataCellRangeAddress);//纵轴为各个数据
        XDDFLineChartData.Series series1 = (XDDFLineChartData.Series) leftdata.addSeries(xs, ys1);
        series1.setTitle("交易量总计（笔）", null); // https://stackoverflow.com/questions/21855842
        series1.setSmooth(false); // https://stackoverflow.com/questions/29014848
        series1.setMarkerStyle(MarkerStyle.DASH); // https://stackoverflow.com/questions/39636138

        CellRangeAddress dataCellRangeAddress2 = new CellRangeAddress(prostartrow, proendrow, prostartcol, proendcol);
        XDDFNumericalDataSource<Double> ys2 = XDDFDataSourcesFactory.fromNumericCellRange(sheet, dataCellRangeAddress2);//纵轴为各个数据
        XDDFLineChartData.Series series2 = (XDDFLineChartData.Series) leftdata.addSeries(xs, ys2);
        series2.setTitle("收入总计（元）", null); // https://stackoverflow.com/questions/21855842
        series2.setSmooth(false); // https://stackoverflow.com/questions/29014848
        series2.setMarkerStyle(MarkerStyle.DASH); // https://stackoverflow.com/questions/39636138

        if (null != paramMap.get("numstartrow2") && null != paramMap.get("numendrow2") && null != paramMap.get("prostartrow2") && null != paramMap.get("proendrow2")) {
            int numstartrow2 = paramMap.get("numstartrow2"), numendrow2 = paramMap.get("numendrow2"),
                    prostartrow2 = paramMap.get("prostartrow2"), proendrow2 = paramMap.get("proendrow2");
            series1.setTitle("类型1交易量总计（笔）", null); // https://stackoverflow.com/questions/21855842
            series2.setTitle("类型1收入总计（元）", null); // https://stackoverflow.com/questions/21855842
            CellRangeAddress dataCellRangeAddress3 = new CellRangeAddress(numstartrow2, numendrow2, numstartcol, numendcol);
            XDDFNumericalDataSource<Double> ys3 = XDDFDataSourcesFactory.fromNumericCellRange(sheet, dataCellRangeAddress3);//纵轴为各个数据
            XDDFLineChartData.Series series3 = (XDDFLineChartData.Series) leftdata.addSeries(xs, ys3);
            series3.setTitle("类型2交易量总计（笔）", null); // https://stackoverflow.com/questions/21855842
            series3.setSmooth(false); // https://stackoverflow.com/questions/29014848
            series3.setMarkerStyle(MarkerStyle.DASH); // https://stackoverflow.com/questions/39636138

            CellRangeAddress dataCellRangeAddress4 = new CellRangeAddress(prostartrow2, proendrow2, prostartcol, proendcol);
            XDDFNumericalDataSource<Double> ys4 = XDDFDataSourcesFactory.fromNumericCellRange(sheet, dataCellRangeAddress4);//纵轴为各个数据
            XDDFLineChartData.Series series4 = (XDDFLineChartData.Series) leftdata.addSeries(xs, ys4);
            series4.setTitle("类型2收入总计（元）", null); // https://stackoverflow.com/questions/21855842
            series4.setSmooth(false); // https://stackoverflow.com/questions/29014848
            series4.setMarkerStyle(MarkerStyle.DASH); // https://stackoverflow.com/questions/39636138
            chart.plot(leftdata);
            chart.plot(leftdata);
        }
        chart.plot(leftdata);
        chart.plot(leftdata);
    }

    /**
     * 柱状图
     *
     * @param sheet
     * @param desc
     * @param paramMap
     */

    private static void drawBarChart(XSSFSheet sheet, String desc, Map<String, Integer> paramMap) {
        int col1 = paramMap.get("col1"), col2 = paramMap.get("col2"), row1 = paramMap.get("row1"), row2 = paramMap.get("row2"),//绘图所在坐标，默认宽度为12列
                numstartrow = paramMap.get("numstartrow"), numendrow = paramMap.get("numendrow"), numstartcol = paramMap.get("numstartcol"), numendcol = paramMap.get("numendcol"),
                prostartrow = paramMap.get("prostartrow"), proendrow = paramMap.get("proendrow"), prostartcol = paramMap.get("prostartcol"), proendcol = paramMap.get("proendcol");

        int dx1 = 0;
        int dy1 = 0;
        int dx2 = 0;
        int dy2 = 0;
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(dx1, dy1, dx2, dy2, col1, row1, col2, row2);

        XSSFChart chart = drawing.createChart(anchor);
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.TOP_RIGHT);

        // Use a category axis for the bottom axis.
        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.TOP);//底部X轴
        bottomAxis.setTitle(desc + "交易情况汇总"); // https://stackoverflow.com/questions/32010765

        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);//左侧Y轴
        leftAxis.setTitle("交易量/金额");
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

        XDDFBarChartData data = (XDDFBarChartData) chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
        CellRangeAddress cellRangeAddress = new CellRangeAddress(numstartrow, numendrow, numstartcol, numendcol);
        XDDFDataSource<String> xs = XDDFDataSourcesFactory.fromStringCellRange(sheet, cellRangeAddress);//横轴为第一列日期

        CellRangeAddress dataCellRangeAddress = new CellRangeAddress(numstartrow, numendrow, numstartcol, numendcol);
        XDDFNumericalDataSource<Double> ys1 = XDDFDataSourcesFactory.fromNumericCellRange(sheet, dataCellRangeAddress);//纵轴为各个数据
        XDDFBarChartData.Series series1 = (XDDFBarChartData.Series) data.addSeries(xs, ys1);
        series1.setTitle("交易量总计（笔）", null); // https://stackoverflow.com/questions/21855842

        CellRangeAddress dataCellRangeAddress2 = new CellRangeAddress(prostartrow, proendrow, prostartcol, proendcol);
        XDDFNumericalDataSource<Double> ys2 = XDDFDataSourcesFactory.fromNumericCellRange(sheet, dataCellRangeAddress2);//纵轴为各个数据
        XDDFBarChartData.Series series2 = (XDDFBarChartData.Series) data.addSeries(xs, ys2);
        series2.setTitle("收入总计（元）", null); // https://stackoverflow.com/questions/21855842
/*
            for(int col=1;col<2;col++) {//数据列数：第一列为日期，其他列为对应数据。
    			CellRangeAddress dataCellRangeAddress=new CellRangeAddress(prostartrow, proendrow, prostartcol, proendcol);
	            XDDFNumericalDataSource<Double> ys = XDDFDataSourcesFactory.fromNumericCellRange(sheet,dataCellRangeAddress);//纵轴为各个数据
	            XDDFChartData.Series series1 = data.addSeries(xs, ys);
	            series1.setTitle("111", null);
            }*/
        chart.plot(data);
        // in order to transform a bar chart into a column chart, you just need to change the bar direction
        XDDFBarChartData bar = (XDDFBarChartData) data;
        bar.setBarDirection(BarDirection.COL);
    }

    /**
     * 隐藏列
     *
     * @param sheet
     * @param sheetName
     */
    private static void hiddenColumn(XSSFSheet sheet, String sheetName) {
		/*if(sheetName.equals("sheetname2")){//需要隐藏列
			sheet.setColumnHidden(3, true);
			sheet.setColumnHidden(4, true);
			sheet.setColumnHidden(5, true);
			sheet.setColumnHidden(6, true);
			sheet.setColumnHidden(7, true);
			sheet.setColumnHidden(8, true);
		}*/
    }

    /**
     * 设置边框
     */
    private static XSSFCellStyle setBorder(XSSFWorkbook wb) {
        XSSFCellStyle style = wb.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);    //边框
        style.setBorderTop(BorderStyle.THIN);        //边框
        style.setBorderLeft(BorderStyle.THIN);        //边框
        style.setBorderRight(BorderStyle.THIN);        //边框
        return style;
    }

}
