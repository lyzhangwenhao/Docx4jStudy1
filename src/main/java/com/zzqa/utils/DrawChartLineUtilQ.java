package com.zzqa.utils;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.axis.DateAxis;
import org.jfree.chart.axis.NumberAxis;
import org.jfree.chart.plot.XYPlot;
import org.jfree.chart.renderer.xy.XYItemRenderer;
import org.jfree.data.xy.XYDataset;
import org.jfree.data.xy.XYSeries;
import org.jfree.data.xy.XYSeriesCollection;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;

/**
 * ClassName: DrawChartLine
 * Description:
 *
 * @author 张文豪
 * @date 2020/8/6 9:24
 */
public class DrawChartLineUtilQ {

    /**
     * 传入数据生成趋势图
     * @param title 机组名称：一般为机组+测点
     * @param xUnit x轴单位
     * @param yUnit y轴单位
     * @param rowKey 折线所属类别
     * @param colKeys 折线x轴数据
     * @param data 与x轴对应的y轴数据
     * @return 返回一个File对象
     */
    public static File getImageFile(String title,String xUnit,String yUnit,
                                    String rowKey,long[] colKeys,double[][] data){
        ChartUtil.setChartTheme();
        // 步骤1：创建CategoryDataset对象（准备数据）
        String[]rowKeys={rowKey};
//        CategoryDataset dataset = DatasetUtilities.createCategoryDataset(rowKeys, colKeys, data);
        XYDataset dataset = createDataset(rowKey,colKeys,data);
        // 步骤2：根据Dataset 生成JFreeChart对象，以及做相应的设置
        JFreeChart freeChart = createChart(dataset,title,xUnit,yUnit);
        //步骤3：将JFreeChart对象保存为一个文件
        BufferedImage bufferedImage = freeChart.createBufferedImage(800, 500);
        //确定图表文件存在的位置，如果文件夹路径不存在则创建一个
        File mkdirPath = new File("D:/chart/lineChart");
        if(!mkdirPath.exists() && !mkdirPath.isDirectory()){
            mkdirPath.mkdirs();
        }

        //折线图路径
        String lineFilePath = mkdirPath.getAbsolutePath()+"/"+System.currentTimeMillis() +title+".png";
        File file = new File(lineFilePath);
        try {
            ImageIO.write(bufferedImage, "png",file);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return file;
    }

    // 根据CategoryDataset创建JFreeChart对象

    /**
     * 根据CategoryDataset创建JFreeChart对象
     * @param xyDataset 生成的CategoryDataset对象
     * @param title 机组+测点
     * @param xUnit x轴单位
     * @param yUnit y轴单位
     * @return 返回一个JFreeChart对象
     */
    private static JFreeChart createChart(XYDataset xyDataset, String title, String xUnit, String yUnit) {
        // 创建JFreeChart对象：ChartFactory.createLineChart
        JFreeChart jfreechart = ChartFactory.createTimeSeriesChart(title, // 标题
                xUnit, // categoryAxisLabel （category轴，横轴，X轴标签）
                yUnit, // valueAxisLabel（value轴，纵轴，Y轴的标签）
                xyDataset, // dataset
                true, // legend
                false, // tooltips
                false); // URLs
        // 使用CategoryPlot设置各种参数。以下设置可以省略。
        XYPlot plot = (XYPlot)jfreechart.getPlot();
        // 背景色 透明度
        plot.setBackgroundAlpha(0.5f);
        // 前景色 透明度
        plot.setForegroundAlpha(0.5f);

        DateAxis domainAxis = (DateAxis)plot.getDomainAxis();
        domainAxis.setPositiveArrowVisible(true);
        domainAxis.setDateFormatOverride(new SimpleDateFormat("yyyy-MM-dd"));
        domainAxis.setMinorTickCount(8);

        // 设置Y轴
        NumberAxis rangeAxis = (NumberAxis) plot.getRangeAxis();
        rangeAxis.setLabelFont(new Font("宋书", Font.PLAIN, 15));
        //纵轴的箭头
        rangeAxis.setPositiveArrowVisible(true);
//        rangeAxis.setUpperMargin(0.0);//上边距,防止最大的一个数据靠近了坐标轴。

        XYItemRenderer renderer = plot.getRenderer();
        //设置线条颜色
        renderer.setSeriesPaint(0, Color.BLUE);

        return jfreechart;
    }

    /**
     * 创建XYDataset对象
     *
     */
    public static XYDataset createDataset(String rowKey,long[] colKeys,double[][] data) {
        XYSeries first = new XYSeries(rowKey);
        for (int i=0;i<colKeys.length;i++){
            first.add(colKeys[i],data[0][i]);
        }
        XYSeriesCollection xySeriesCollection = new XYSeriesCollection();
        xySeriesCollection.addSeries(first);
        return xySeriesCollection;
    }
}
