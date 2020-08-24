package com.zzqa.docx4j2word;

import com.zzqa.pojo.*;
import com.zzqa.utils.Docx4jUtil;
import com.zzqa.utils.LoadDataUtils;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.toc.Toc;
import org.docx4j.toc.TocGenerator;
import org.docx4j.toc.TocHelper;

import java.io.File;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Random;

/**
 * ClassName: Docx4j2WordMain
 * Description:
 *
 * @author 张文豪
 * @date 2020/8/3 11:17
 */
public class Docx4j2WordMain {
    public static void main(String[] args) {
        try {
            WordprocessingMLPackage wpMLPackage = WordprocessingMLPackage.createPackage();

            //数据准备
            String reportName = "风电场风电机组";  //项目名称
            long startTime = 1468166400000L;    //报告开始时间毫秒值
            long endTime = System.currentTimeMillis();  //报告结束时间毫秒值
            String logoPath = "D:\\01_ideaspace\\Docx4jStudy\\src\\main\\resources\\images\\logo.png"; //封面logo路径
            String linePath = "src\\main\\resources\\images\\横线.png";   //封面横线路径

            int normalPart = 106;   //正常机组数量
            int warningPart = 20;   //预警机组数量
            int alarmPart = 12; //报警机组数量
            //最后保存的文件名
            String fileName = reportName + getDate(new Date(startTime)) + getDate(new Date(endTime)) + "检测报告.docx";
            //保存的文件路径
            String targetFilePath = "D:/AutoExport";
            String targetFile = targetFilePath + "/" + fileName;
            //PageContent2的数据准备
            List<Characteristic> characteristicList = createData();
            List<List<Waveform>> chartData = createChartData();


            Cover cover = new Cover();
            //创建封面
            wpMLPackage = cover.createCover(wpMLPackage, reportName, startTime, endTime, logoPath, linePath);
            //在这个前面添加目录
            Docx4jUtil.addNextSection(wpMLPackage);
            //文件内容1:项目概述
            PageContent1 pageContent1 = new PageContent1();
            pageContent1.createPageContent1(wpMLPackage, reportName, normalPart, warningPart, alarmPart);
            //文件内容2:运行状况
            PageContent2 pageContent2 = new PageContent2();
            pageContent2.createPageContent2(wpMLPackage, characteristicList);
            //文件内容3：震动图谱
            PageContent3 pageContent3 = new PageContent3();
            pageContent3.createPageContent3(wpMLPackage, chartData, 0);
            //文件内容4：补充说明
            PageContent4 pageContent4 = new PageContent4();
            pageContent4.createPageContent4(wpMLPackage);

            //保存文件
            File docxFile = new File(targetFilePath);
            if (!docxFile.exists() && !docxFile.isDirectory()) {
                docxFile.mkdirs();
            }

            //生成目录
            TocGenerator tocGenerator = new TocGenerator(wpMLPackage);
            Toc.setTocHeadingText("目录");
            tocGenerator.generateToc(15, TocHelper.DEFAULT_TOC_INSTRUCTION, true);

            wpMLPackage.save(new File(targetFile));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 数据准备
     *
     * @return
     */
    public static List<Characteristic> createData() {
        String[] codeNameArr = {"有效值","峰值","峰峰值"};
        Random random = new Random();
        //数据集合
        List<Characteristic> characteristicList = new ArrayList<>();
        //循环向集合中添加
        for (int i1=0; i1<10; i1++){
            List<Characteristic> characteristicList1 = new ArrayList<>();
            Characteristic characteristic1 = new Characteristic();
            //生成机组名称
            characteristic1.setName("机组#"+i1);
            for (int i2=0; i2<5; i2++){
                List<Characteristic> characteristicList2 = new ArrayList<>();
                Characteristic characteristic2 = new Characteristic();
                //生成部件名称
                characteristic2.setName("部件#"+i2);
                for (int i3=0; i3<3; i3++){
                    Characteristic characteristic3 = new Characteristic();
                    List<Feature> featureList = new ArrayList<>();
                    //生成测点名称
                    characteristic3.setName("测点#"+i3);
                    //特征值数据

                    for (int i4=0; i4<3; i4++){
                        Feature feature = new Feature();
                        feature.setCodeName(codeNameArr[random.nextInt(3)]);
                        feature.setCount(i4);
                        feature.setMaxValue(String.valueOf(i4));
                        feature.setValve(i4+"g");
                        feature.setLevel(random.nextInt(2)==0?"报警":"预警");
                        featureList.add(feature);
                    }
                    characteristic3.setF_list(featureList);
                    characteristicList2.add(characteristic3);
                }
                characteristic2.setList(characteristicList2);
                characteristicList1.add(characteristic2);
            }
            characteristic1.setList(characteristicList1);
            characteristicList.add(characteristic1);
        }

        return characteristicList;
    }


    public static List<List<Waveform>> createChartData(){
        Random random = new Random();
        List<List<Waveform>> list = new ArrayList<>();
        for (int j=0; j<5; j++){
            List<Waveform> list1 = new ArrayList<>();
            for (int i=0; i<6; i++){
                String machineName = "机组"+(j+1);
                String positionName = "测点"+(i+1);
                String level = random.nextInt(2)==1?"报警":"预警";
                Waveform waveform = createWaveform(machineName, positionName, level);
                list1.add(waveform);
            }
            list.add(list1);
        }
        return list;
    }

    public static Waveform createWaveform(String machineName,String positionName, String level){
        Waveform waveform = new Waveform();
        //趋势图
        String dataX1 = LoadDataUtils.ReadFile("C:/Users/Mi_dad/Desktop/趋势图X.txt");
        String[] colKeys = dataX1.split(",");
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat();
        simpleDateFormat.applyPattern("yyyy/MM/dd HH:mm");
        long [] colKeys1 = new long[colKeys.length];    //趋势图X
        for (int i=0; i<colKeys.length; i++){
            try {
                Date parse = simpleDateFormat.parse(colKeys[i]);
                colKeys1[i] = parse.getTime();
            } catch (ParseException e) {
                e.printStackTrace();
            }
        }
        String dataY1 = LoadDataUtils.ReadFile("C:/Users/Mi_dad/Desktop/趋势图Y.txt");
        double[][] data1 = string2DoubleArray(dataY1);  //趋势图Y
        //波形图
        String dataX2 = LoadDataUtils.ReadFile("C:/Users/Mi_dad/Desktop/波形图X.txt");
        String[] colKeys2 = dataX2.split(",");  //波形图X
        String dataY2 = LoadDataUtils.ReadFile("C:/Users/Mi_dad/Desktop/波形图Y.txt");
        double[][] data2 = string2DoubleArray(dataY2);  //波形图Y
        //频谱图
        String dataX3 = LoadDataUtils.ReadFile("C:/Users/Mi_dad/Desktop/频谱图X.txt");
        String[] colKeys3 = dataX3.split(",");  //频谱图X
        String dataY3 = LoadDataUtils.ReadFile("C:/Users/Mi_dad/Desktop/频谱图Y.txt");
        double[][] data3 = string2DoubleArray(dataY3);  //频谱图Y
        //包络图
        String dataX4 = LoadDataUtils.ReadFile("C:/Users/Mi_dad/Desktop/包络图X.txt");
        String[] colKeys4 = dataX4.split(",");  //包络图X
        String dataY4 = LoadDataUtils.ReadFile("C:/Users/Mi_dad/Desktop/包络图Y.txt");
        double[][] data4 = string2DoubleArray(dataY4);  //包络图Y


        //趋势图集合
        WaveformChild waveformChild11 = new WaveformChild();
        waveformChild11.setFeatureName("峰值");
        waveformChild11.setValue_x(colKeys1);
        waveformChild11.setValue(data1);

        WaveformChild waveformChild12 = new WaveformChild();
        waveformChild12.setFeatureName("有效值");
        waveformChild12.setValue_x(colKeys1);
        waveformChild12.setValue(data1);

        List<WaveformChild> waveformChildList = new ArrayList<>();
        waveformChildList.add(waveformChild11);
        waveformChildList.add(waveformChild12);
        //波形图
        WaveformChild waveformChild2 = new WaveformChild();
        waveformChild2.setWave_x(string2Double(colKeys2));
        waveformChild2.setValue(data2);
        //频谱图
        WaveformChild waveformChild3 = new WaveformChild();
        waveformChild3.setWave_x(string2Double(colKeys3));
        waveformChild3.setValue(data3);
        //包络图
        WaveformChild waveformChild4 = new WaveformChild();
        waveformChild4.setWave_x(string2Double(colKeys4));
        waveformChild4.setValue(data4);

        waveform.setMachineName(machineName);
        waveform.setPositionName(positionName);
        waveform.setLevel(level);
        waveform.setWave(waveformChild2);
        waveform.setSpectrum(waveformChild3);
        waveform.setSpm(waveformChild4);
        waveform.setList(waveformChildList);

        return waveform;
    }
    /**
     * 将数据转换为double数组
     *
     * @param dataY
     * @return
     */
    private static double[][] string2DoubleArray(String dataY) {
        String[] split = dataY.split(",");
        double[] dataTemp = new double[split.length];
        for (int i = 0; i < split.length; i++) {
            if (split[i] != null) {
                dataTemp[i] = Double.parseDouble(split[i]);
            }
        }
        double[][] data = {dataTemp};
        return data;
    }

    /**
     * String类型转换为Double类型
     * @param colKeys
     * @return
     */
    public static double[] string2Double(String[] colKeys){
        double [] colKeys1 = new double[colKeys.length];
        for (int i=0; i<colKeys.length; i++){
            colKeys1[i] = Double.parseDouble(colKeys[i]);
        }
        return colKeys1;
    }

    public static String getDate(Date date) {
        SimpleDateFormat sdf = new SimpleDateFormat();// 格式化时间
        sdf.applyPattern("yyyy年M月d日");
        return sdf.format(date);
    }
}
