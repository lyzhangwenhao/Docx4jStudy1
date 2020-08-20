package com.zzqa.pojo;

/**
 * @author wangguocai
 * @date 2020-08-19 11:10
 */
public class Feature {
    private int count;  //报警次数
    private String valve;   //触发门限
    private String codeName;    //特征值
    private String maxValue;    //最大报警特征值
    private String level;   //报警等级

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

    public String getCodeName() {
        return codeName;
    }

    public void setCodeName(String codeName) {
        this.codeName = codeName;
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
}
