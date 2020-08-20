package com.zzqa.pojo;

/**
 * @author wangguocai
 * @date 2020-08-19 13:50
 */
public class WaveformChild {
    private String featureName;//特征值名称
    private long[] value_x;//趋势图x轴
    private double[]  wave_x;//波形图、频谱图、包络图的x轴
    private double[][] value;//波数据

    public long[] getValue_x() {
        return value_x;
    }

    public void setValue_x(long[] value_x) {
        this.value_x = value_x;
    }

    public String getFeatureName() {
        return featureName;
    }

    public void setFeatureName(String featureName) {
        this.featureName = featureName;
    }

    public double[][] getValue() {
        return value;
    }

    public void setValue(double[][] value) {
        this.value = value;
    }

    public double[] getWave_x() {
        return wave_x;
    }

    public void setWave_x(double[] wave_x) {
        this.wave_x = wave_x;
    }
}
