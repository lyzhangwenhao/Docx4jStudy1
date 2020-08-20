package com.zzqa.docx4j2word;

import com.zzqa.pojo.UnitInfo;
import com.zzqa.utils.Docx4jUtil;
import com.zzqa.utils.DrawChartLineUtil;
import com.zzqa.utils.DrawChartLineUtilQ;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.ObjectFactory;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

/**
 * ClassName: PageConten3
 * Description:
 *
 * @author 张文豪
 * @date 2020/8/6 14:48
 */
public class PageContent3 {
    private ObjectFactory factory = new ObjectFactory();

    /**
     * 生成预警/报警机组的四种图谱
     *
     * @param wpMLPackage 传入的wpMLPackage对象
     * @param unitInfos   数据
     * @param mac_export  正常机组是否导出，1为导出
     */
    public void createPageContent3(WordprocessingMLPackage wpMLPackage, List<List<List<List<UnitInfo>>>> unitInfos, int mac_export) {
        try {
            //添加标题三：震动图谱
            wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading1", "3 震动图谱");
            //三种报警等级
            List<UnitInfo> unitInfos1 = new ArrayList<>();  //报警机组数据
            List<UnitInfo> unitInfos2 = new ArrayList<>();  //预警机组数据
            List<UnitInfo> unitInfos3 = new ArrayList<>();  //正常机组数据
            //将数据分离，分别分为报警，预警，正常
            if (unitInfos != null) {
                for (List<List<List<UnitInfo>>> list1 : unitInfos) {
                    for (List<List<UnitInfo>> list2 : list1) {
                        for (List<UnitInfo> list3 : list2) {
                            for (UnitInfo u : list3) {
                                if ("报警".equals(u.getLevel())) {
                                    unitInfos1.add(u);
                                } else if ("预警".equals(u.getLevel())) {
                                    unitInfos2.add(u);
                                } else {
                                    unitInfos3.add(u);
                                }
                            }
                        }
                    }
                }
            }
            //添加四种图谱
            //注意，1、图表数据必须是根据X轴排好序的，(x,y)一一对应，不然图像不对
            //      2、y轴的数据必须为double类型，所以，字符串中不能包含任何不能转换为double类型的数据
            //报警
            addImageContent(wpMLPackage, unitInfos1, 1);
            //预警
            addImageContent(wpMLPackage, unitInfos2, 2);
            //正常
            if (mac_export == 1) {
                addImageContent(wpMLPackage, unitInfos3, 3);
            }
            //TODO 删除输出语句
            System.out.println("PageContent3 Success......");

        } catch (Exception e) {
            e.printStackTrace();
        }

        //4、包络图

        //下一页
        Docx4jUtil.addNextPage(wpMLPackage);
    }

    /**
     * 添加四种图谱到wpMLPackage中
     *
     * @param wpMLPackage
     * @param unitInfos
     * @param flag        标记预警等级
     * @throws Exception
     */
    private void addImageContent(WordprocessingMLPackage wpMLPackage, List<UnitInfo> unitInfos, int flag) throws Exception {
        String index = "3.1";
        String title = "报警机组";
        if (flag == 1) {
            index = "3.1";
            title = "报警机组";
        } else if (flag == 2) {
            index = "3.2";
            title = "预警机组";
        } else if (flag == 3) {
            index = "3.3";
            title = "正常机组";
        }
        if (unitInfos != null && unitInfos.size() != 0) {
            wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading2", index + " " + title);

            Map<String, List<UnitInfo>> collect = unitInfos.stream().collect(Collectors.groupingBy(u -> u.getUnitName() + u.getUnitPart() + u.getUnitStation()));
            AtomicInteger titleIndex = new AtomicInteger(0); //标题序号
            String finalIndex = index;  //如果不添加这个中间变量stream流中使用会报错
            collect.entrySet().stream().sorted(Map.Entry.comparingByKey()).forEachOrdered(map -> {

                String key = map.getKey();
                List<UnitInfo> value = map.getValue();
                titleIndex.getAndIncrement();
                //添加三级标题
                wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading3", finalIndex + "." + titleIndex + " " + key);
                for (int i = 0; i < value.size(); i++) {
                    UnitInfo unitInfo = value.get(i);
                    //1、趋势图
                    File imageFile = DrawChartLineUtilQ.getImageFile(unitInfo.getUnitName() + unitInfo.getUnitPart() + unitInfo.getUnitStation(),
                            "t(s)", "g", unitInfo.getType() + "趋势图线", unitInfo.getColKeys1(), unitInfo.getData1());
                    byte[] bytes = Docx4jUtil.convertImageToByteArray(imageFile);
                    Docx4jUtil.addImageToPackage(wpMLPackage, bytes);
                    deleteImageFile(imageFile);
                    Docx4jUtil.addTableTitle(wpMLPackage,
                            "图" + finalIndex + "." + titleIndex + "-" + (i + 1) + " " +
                                    unitInfo.getUnitName() + "-" + unitInfo.getUnitPart() + "-" + unitInfo.getUnitStation() + unitInfo.getType() + "趋势图");
                    if (i == value.size() - 1) {
                        //2、波形图
                        File imageFile2 = DrawChartLineUtil.getImageFile(unitInfo.getUnitName() + unitInfo.getUnitPart() + unitInfo.getUnitStation(),
                                "t(s)", "g", "波线图线", unitInfo.getColKeys2(), unitInfo.getData2());
                        byte[] bytes2 = Docx4jUtil.convertImageToByteArray(imageFile2);
                        Docx4jUtil.addImageToPackage(wpMLPackage, bytes2);
                        deleteImageFile(imageFile2);
                        Docx4jUtil.addTableTitle(wpMLPackage,
                                "图" + finalIndex + "." + titleIndex + "-" + (i + 1) + " " +
                                        unitInfo.getUnitName() + "-" + unitInfo.getUnitPart() + "-" + unitInfo.getUnitStation() + "波形图");

                        //3、频谱图
                        File imageFile3 = DrawChartLineUtil.getImageFile(unitInfo.getUnitName() + unitInfo.getUnitPart() + unitInfo.getUnitStation(),
                                "t(s)", "g", "频谱图线", unitInfo.getColKeys3(), unitInfo.getData3());
                        byte[] bytes3 = Docx4jUtil.convertImageToByteArray(imageFile3);
                        Docx4jUtil.addImageToPackage(wpMLPackage, bytes3);
                        deleteImageFile(imageFile3);
                        Docx4jUtil.addTableTitle(wpMLPackage,
                                "图" + finalIndex + "." + titleIndex + "-" + (i + 1) + " " +
                                        unitInfo.getUnitName() + "-" + unitInfo.getUnitPart() + "-" + unitInfo.getUnitStation() + "频谱图");

                        //4、包络图
                        File imageFile4 = DrawChartLineUtil.getImageFile(unitInfo.getUnitName() + unitInfo.getUnitPart() + unitInfo.getUnitStation(),
                                "t(s)", "g", "包络图线", unitInfo.getColKeys4(), unitInfo.getData4());
                        byte[] bytes4 = Docx4jUtil.convertImageToByteArray(imageFile4);
                        Docx4jUtil.addImageToPackage(wpMLPackage, bytes4);
                        deleteImageFile(imageFile4);
                        Docx4jUtil.addTableTitle(wpMLPackage,
                                "图" + finalIndex + "." + titleIndex + "-" + (i + 1) + " " +
                                        unitInfo.getUnitName() + "-" + unitInfo.getUnitPart() + "-" + unitInfo.getUnitStation() + "包络图");
                    }

                }


            });
        }


//        if (unitInfos != null && unitInfos.size() != 0) {
//            wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading2", index + " " + title);
//            UnitInfo unitInfo = new UnitInfo();
//
//
//            for (int i = 0; i < unitInfos.size(); i++) {
//                //图表索引序号
//                int num = 1;
//                unitInfo = unitInfos.get(i);
//                wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading3", index + "." + (i + 1) + " " + unitInfo.getUnitName() + "-" + unitInfo.getUnitPart() + "-" + unitInfo.getUnitStation());
//
//
//                //1、趋势图
//                File imageFile = DrawChartLineUtilQ.getImageFile(unitInfo.getUnitName() + unitInfo.getUnitPart() + unitInfo.getUnitStation(),
//                        "t(s)", "g", "趋势图线", unitInfo.getColKeys1(), unitInfo.getData1());
//                byte[] bytes = Docx4jUtil.convertImageToByteArray(imageFile);
//                Docx4jUtil.addImageToPackage(wpMLPackage, bytes);
//                deleteImageFile(imageFile);
//                Docx4jUtil.addTableTitle(wpMLPackage,
//                        "图" + index + "." + (i + 1) + "-" + num + " " +
//                                unitInfo.getUnitName() + "-" + unitInfo.getUnitPart() + "趋势图");
//                num++;
//                //2、波形图
//                File imageFile2 = DrawChartLineUtil.getImageFile(unitInfo.getUnitName() + unitInfo.getUnitPart() + unitInfo.getUnitStation(),
//                        "t(s)", "g", "波线图线", unitInfo.getColKeys2(), unitInfo.getData2());
//                byte[] bytes2 = Docx4jUtil.convertImageToByteArray(imageFile2);
//                Docx4jUtil.addImageToPackage(wpMLPackage, bytes2);
//                deleteImageFile(imageFile2);
//                Docx4jUtil.addTableTitle(wpMLPackage,
//                        "图" + index + "." + (i + 1) + "-" + num + " " +
//                                unitInfo.getUnitName() + "-" + unitInfo.getUnitPart() + "波形图");
//                num++;
//
//                //3、频谱图
//                File imageFile3 = DrawChartLineUtil.getImageFile(unitInfo.getUnitName() + unitInfo.getUnitPart() + unitInfo.getUnitStation(),
//                        "t(s)", "g", "频谱图线", unitInfo.getColKeys3(), unitInfo.getData3());
//                byte[] bytes3 = Docx4jUtil.convertImageToByteArray(imageFile3);
//                Docx4jUtil.addImageToPackage(wpMLPackage, bytes3);
//                deleteImageFile(imageFile3);
//                Docx4jUtil.addTableTitle(wpMLPackage,
//                        "图" + index + "." + (i + 1) + "-" + num + " " +
//                                unitInfo.getUnitName() + "-" + unitInfo.getUnitPart() + "频谱图");
//                num++;
//
//                //4、包络图
//                File imageFile4 = DrawChartLineUtil.getImageFile(unitInfo.getUnitName() + unitInfo.getUnitPart() + unitInfo.getUnitStation(),
//                        "t(s)", "g", "包络图线", unitInfo.getColKeys4(), unitInfo.getData4());
//                byte[] bytes4 = Docx4jUtil.convertImageToByteArray(imageFile4);
//                Docx4jUtil.addImageToPackage(wpMLPackage, bytes4);
//                deleteImageFile(imageFile4);
//                Docx4jUtil.addTableTitle(wpMLPackage,
//                        "图" + index + "." + (i + 1) + "-" + num + " " +
//                                unitInfo.getUnitName() + "-" + unitInfo.getUnitPart() + "包络图");
//                num++;
//
//            }
//        }
    }

    /**
     * 删除已经添加进文档中的图片
     *
     * @param imageFile
     */
    private void deleteImageFile(File imageFile) {
        if (imageFile != null && imageFile.exists()) {
            imageFile.delete();
        }
    }

    /**
     * 将数据转换为double数组
     *
     * @param dataY
     * @return
     */
    private double[][] string2DoubleArray(String dataY) {
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
}
