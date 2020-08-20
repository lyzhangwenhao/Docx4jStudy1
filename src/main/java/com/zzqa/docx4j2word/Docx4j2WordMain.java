package com.zzqa.docx4j2word;

import com.zzqa.pojo.Characteristic;
import com.zzqa.pojo.Feature;
import com.zzqa.pojo.UnitInfo;
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
            //List<Characteristic> characteristicList
            List<Characteristic> characteristicList = createData();
            //TODO 删除下面输出
//            characteristicList.stream().forEach(l->{
//                l.stream().forEach(s->{
//                    s.stream().forEach(p->{
//                        System.out.println(p.toString());
//                    });
//                });
//            });


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
//            pageContent3.createPageContent3(wpMLPackage, characteristicList, 0);
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
        //TODO 删除输出语句
        System.out.println("转换时间：" + sdf.format(date)); // 输出已经格式化的现在时间（24小时制）
        return sdf.format(date);
    }
}
