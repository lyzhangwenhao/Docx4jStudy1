package com.zzqa.utils;

import org.jfree.chart.ChartColor;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.labels.StandardPieSectionLabelGenerator;
import org.jfree.chart.plot.PiePlot;
import org.jfree.chart.title.TextTitle;
import org.jfree.data.general.DefaultPieDataset;
import org.jfree.data.general.PieDataset;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.NumberFormat;

/**
 * ClassName: DrawChartPieUtil
 * Description:
 *
 * @author 张文豪
 * @date 2020/8/4 16:39
 */
public class DrawChartPieUtil {
    /**
     * 根据传入的数据获取图片的File文件
     * @param proName 项目名称
     * @param normalPart 正常机组数量
     * @param warningPart 预警机组数量
     * @param alarmPart 警告机组数量
     * @return 返回File文件类型
     */
    public static File getImageFile(String proName,int normalPart,int warningPart,int alarmPart)
    {
        //设置数据源
        PieDataset mDataset = GetDataset(normalPart,warningPart,alarmPart);
        //建立图表
        JFreeChart mChart = ChartFactory.createPieChart(proName,mDataset);

        //设置图表标题
        mChart.setTitle(new TextTitle(proName, new Font("黑体",Font.CENTER_BASELINE, 20)));
        //设置Legend字体
        mChart.getLegend().setItemFont(new Font("宋体", Font.ROMAN_BASELINE, 15));

        PiePlot mPiePlot = (PiePlot) mChart.getPlot();
        //设置背景色
        mPiePlot.setBackgroundPaint(ChartColor.WHITE);
        //数据区边框隐藏
        mPiePlot.setOutlineVisible(false);

        mPiePlot.setLabelGenerator(new StandardPieSectionLabelGenerator("{0}={1}({2})", NumberFormat.getNumberInstance(), new DecimalFormat("0.00%")));
        // 底部图例显示百分比:自定义方式， {0} 表示选项， {1} 表示数值， {2} 表示所占比例
        mPiePlot.setLegendLabelGenerator(new StandardPieSectionLabelGenerator("{0}={1}({2})"));
        //设置饼图标签字体
        mPiePlot.setLabelFont(new Font("宋体", Font.PLAIN, 12));

        //将图标转换为BufferedImage类型
        BufferedImage bufferedImage = mChart.createBufferedImage(475, 300);
        //确定图表文件存在的位置，如果文件夹路径不存在则创建一个
        File mkdirPath = new File("D:/chart/pieChart");
        if(!mkdirPath.exists() && !mkdirPath.isDirectory()){
            mkdirPath.mkdirs();
        }

        //饼状图路径
        String pieFilePath = mkdirPath.getAbsolutePath()+"/"+System.currentTimeMillis() +proName+".png";

        File file = new File(pieFilePath);
        try {
            //保存图表为一个File文件类型
            ImageIO.write(bufferedImage, "png", file);
            //TODO 删除输出语句
            System.out.println("饼状图完成！");
        } catch (IOException e) {
            e.printStackTrace();
        }
        return file;

    }

    /**
     * 设置饼状图的数据
     * @param normalPort 正常机组数量
     * @param warningPort 预警机组数量
     * @param alarmPort 警告机组数量
     * @return 返回一个饼状图数据源
     */
    private static PieDataset GetDataset(int normalPort,int warningPort,int alarmPort)
    {
        DefaultPieDataset mDataset = new DefaultPieDataset();
        mDataset.setValue(" 报警机组", new Double(alarmPort));
        mDataset.setValue(" 预警机组", new Double(warningPort));
        mDataset.setValue(" 正常机组", new Double(normalPort));
        return mDataset;
    }
}
