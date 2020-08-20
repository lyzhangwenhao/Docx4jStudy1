package com.zzqa.pojo;

import java.util.List;

/**
 * @author wangguocai
 * @date 2020-08-19 13:44
 */
public class Waveform {

    private String machineName;//机组名
    private String positionName;//机组下单个测点名称
    private String level;   //报警或预警
    private WaveformChild wave;//波形图
    private WaveformChild spm;//包络图
    private WaveformChild spectrum;//频谱图
    private List<WaveformChild> list;//趋势图

    public String getLevel() {
        return level;
    }

    public void setLevel(String level) {
        this.level = level;
    }

    public String getMachineName() {
        return machineName;
    }

    public void setMachineName(String machineName) {
        this.machineName = machineName;
    }

    public String getPositionName() {
        return positionName;
    }

    public void setPositionName(String positionName) {
        this.positionName = positionName;
    }

    public WaveformChild getWave() {
        return wave;
    }

    public void setWave(WaveformChild wave) {
        this.wave = wave;
    }

    public WaveformChild getSpm() {
        return spm;
    }

    public void setSpm(WaveformChild spm) {
        this.spm = spm;
    }

    public WaveformChild getSpectrum() {
        return spectrum;
    }

    public void setSpectrum(WaveformChild spectrum) {
        this.spectrum = spectrum;
    }

    public List<WaveformChild> getList() {
        return list;
    }

    public void setList(List<WaveformChild> list) {
        this.list = list;
    }
}
