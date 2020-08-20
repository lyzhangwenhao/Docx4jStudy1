package com.zzqa.docx4j2word;

import com.zzqa.pojo.UnitInfo;
import com.zzqa.pojo.Waveform;
import com.zzqa.pojo.WaveformChild;
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
     * @param list        数据
     * @param mac_export  正常机组是否导出，1为导出
     */
    public void createPageContent3(WordprocessingMLPackage wpMLPackage, List<List<Waveform>> list, int mac_export) {
        try {
            //添加标题三：震动图谱
            wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading1", "3 震动图谱");
            //三种报警等级
            //添加四种图谱
            //注意，1、图表数据必须是根据X轴排好序的，(x,y)一一对应，不然图像不对
            //      2、y轴的数据必须为double类型，所以，字符串中不能包含任何不能转换为double类型的数据
            //报警
            addImageContent(wpMLPackage, list, 1);
            //预警
            addImageContent(wpMLPackage, list, 2);
            //正常
            if (mac_export == 1) {
                addImageContent(wpMLPackage, list, 3);
            }
            //TODO 删除输出语句
            System.out.println("PageContent3 Success......");

        } catch (Exception e) {
            e.printStackTrace();
        }

        //下一页
        Docx4jUtil.addNextPage(wpMLPackage);
    }

    /**
     * 添加四种图谱到wpMLPackage中
     *
     * @param wpMLPackage
     * @param list
     * @param flag        标记预警等级
     * @throws Exception
     */
    private void addImageContent(WordprocessingMLPackage wpMLPackage, List<List<Waveform>> list, int flag) throws Exception {
        String index = "3.1";
        String title = "报警机组";
        String level = "";
        if (flag == 1) {
            index = "3.1";
            title = "报警机组";
            level = "报警";
        } else if (flag == 2) {
            index = "3.2";
            title = "预警机组";
            level = "预警";
        } else if (flag == 3) {
            index = "3.3";
            title = "正常机组";
            level = "正常";
        }
        if (list != null && list.size() != 0) {
            wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading2", index + " " + title);
            int positionNum = 1;
            for (List<Waveform> waveforms : list) {
                for (Waveform waveform : waveforms) {
                    if (!level.equals(waveform.getLevel())){
                        continue;
                    }
                    int chartNum = 1;
                    String machineName = waveform.getMachineName();
                    String positionName = waveform.getPositionName();
                    //三级标题
                    String title3 = index + "." + positionNum + " " + machineName + positionName;
                    wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading3", title3);


                    //趋势图
                    List<WaveformChild> waveformList = waveform.getList();
                    for (WaveformChild waveformChild : waveformList) {
                        //1、趋势图
                        long[] value_x = waveformChild.getValue_x();
                        double[][] value_y = waveformChild.getValue();

                        File imageFile = DrawChartLineUtilQ.getImageFile(machineName + positionName,
                                "t(s)", "g", waveformChild.getFeatureName() + "趋势图线", value_x, value_y);
                        byte[] bytes = Docx4jUtil.convertImageToByteArray(imageFile);
                        Docx4jUtil.addImageToPackage(wpMLPackage, bytes);
                        deleteImageFile(imageFile);
                        Docx4jUtil.addTableTitle(wpMLPackage,
                                "图" + index + "." + positionNum + "-" + chartNum + " " +
                                        machineName + "-" + positionName + waveformChild.getFeatureName() + "趋势图");
                        chartNum++;
                    }
                    //波形图
                    //2、波形图
                    WaveformChild wave = waveform.getWave();
                    double[] wave_x = wave.getWave_x();
                    double[][] wave_y = wave.getValue();

                    File imageFile2 = DrawChartLineUtil.getImageFile(machineName + positionName,
                            "t(s)", "g", "波形图线", wave_x, wave_y);
                    byte[] bytes2 = Docx4jUtil.convertImageToByteArray(imageFile2);
                    Docx4jUtil.addImageToPackage(wpMLPackage, bytes2);
                    deleteImageFile(imageFile2);
                    Docx4jUtil.addTableTitle(wpMLPackage,
                            "图" + index + "." + positionNum + "-" + chartNum + " " +
                                    machineName + "-" + positionName + "波形图");
                    chartNum++;
                    //频谱图
                    WaveformChild spectrum = waveform.getSpectrum();
                    double[] spectrumWave_x = spectrum.getWave_x();
                    double[][] spectrumValue = spectrum.getValue();

                    File imageFile3 = DrawChartLineUtil.getImageFile(machineName + positionName,
                            "t(s)", "g", "频谱图线", spectrumWave_x, spectrumValue);
                    byte[] bytes3 = Docx4jUtil.convertImageToByteArray(imageFile3);
                    Docx4jUtil.addImageToPackage(wpMLPackage, bytes3);
                    deleteImageFile(imageFile3);
                    Docx4jUtil.addTableTitle(wpMLPackage,
                            "图" + index + "." + positionNum + "-" + chartNum + " " +
                                    machineName + "-" + positionName + "频谱图");
                    chartNum++;
                    //包络图
                    WaveformChild spm = waveform.getSpm();
                    double[] spmWave_x = spm.getWave_x();
                    double[][] spmValue = spm.getValue();

                    File imageFile4 = DrawChartLineUtil.getImageFile(machineName + positionName,
                            "t(s)", "g", "包络图线", spmWave_x, spmValue);
                    byte[] bytes4 = Docx4jUtil.convertImageToByteArray(imageFile4);
                    Docx4jUtil.addImageToPackage(wpMLPackage, bytes4);
                    deleteImageFile(imageFile4);
                    Docx4jUtil.addTableTitle(wpMLPackage,
                            "图" + index + "." + positionNum + "-" + chartNum + " " +
                                    machineName + "-" + positionName + "包络图");
                    positionNum++;
                }
            }
        }
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
