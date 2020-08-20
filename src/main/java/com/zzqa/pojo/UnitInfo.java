package com.zzqa.pojo;

/**
 * ClassName: UnitInfo
 * Description:
 *
 * @author 张文豪
 * @date 2020/8/5 13:57
 */
public class UnitInfo {
    private String unitName;    //报警机组
    private String unitPart;    //报警部件
    private String unitStation; //报警测点
    private int count;  //报警次数
    private String valve;   //触发门限
    private String type;    //类型
    private String maxValue;    //最大报警特征值
    private String level;   //报警等级

    //趋势图数据
    private long[] colKeys1;
    private double[][] data1;
    //波形图数据
    private double[] colKeys2;
    private double[][] data2;
    //频谱图数据
    private double[] colKeys3;
    private double[][] data3;
    //包络图数据
    private double[] colKeys4;
    private double[][] data4;

    public long[] getColKeys1() {
        return colKeys1;
    }

    public void setColKeys1(long[] colKeys1) {
        this.colKeys1 = colKeys1;
    }

    public double[][] getData1() {
        return data1;
    }

    public void setData1(double[][] data1) {
        this.data1 = data1;
    }

    public double[] getColKeys2() {
        return colKeys2;
    }

    public void setColKeys2(double[] colKeys2) {
        this.colKeys2 = colKeys2;
    }

    public double[] getColKeys3() {
        return colKeys3;
    }

    public void setColKeys3(double[] colKeys3) {
        this.colKeys3 = colKeys3;
    }

    public double[] getColKeys4() {
        return colKeys4;
    }

    public void setColKeys4(double[] colKeys4) {
        this.colKeys4 = colKeys4;
    }

    public String getUnitStation() {
        return unitStation;
    }

    public void setUnitStation(String unitStation) {
        this.unitStation = unitStation;
    }


    public double[][] getData2() {
        return data2;
    }

    public void setData2(double[][] data2) {
        this.data2 = data2;
    }



    public double[][] getData3() {
        return data3;
    }

    public void setData3(double[][] data3) {
        this.data3 = data3;
    }



    public double[][] getData4() {
        return data4;
    }

    public void setData4(double[][] data4) {
        this.data4 = data4;
    }

    public String getUnitName() {
        return unitName;
    }

    public void setUnitName(String unitName) {
        this.unitName = unitName;
    }

    public String getUnitPart() {
        return unitPart;
    }

    public void setUnitPart(String unitPart) {
        this.unitPart = unitPart;
    }

    public int getCount() {
        return count;
    }

    public void setCount(int count) {
        this.count = count;
    }

    public String getValve() {
        return valve;
    }

    public void setValve(String valve) {
        this.valve = valve;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getMaxValue() {
        return maxValue;
    }

    public void setMaxValue(String maxValue) {
        this.maxValue = maxValue;
    }

    public String getLevel() {
        return level;
    }

    public void setLevel(String level) {
        this.level = level;
    }

    @Override
    public String toString() {
        return "UnitInfo{" +
                "unitName='" + unitName + '\'' +
                ", unitPart='" + unitPart + '\'' +
                ", unitStation='" + unitStation + '\'' +
                ", count=" + count +
                ", valve='" + valve + '\'' +
                ", type='" + type + '\'' +
                ", maxValue='" + maxValue + '\'' +
                ", level='" + level + '\'' +
                '}';
    }
}
