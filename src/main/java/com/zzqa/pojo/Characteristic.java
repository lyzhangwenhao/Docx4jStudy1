package com.zzqa.pojo;

import java.util.List;

/**
 * @author wangguocai
 * @date 2020-08-19 10:33
 */
public class Characteristic {
    private String name;//机组、部件、测点名称
    private List<Characteristic> list;
    private List<Feature> f_list;//测点下特征值内容集合

    public List<Feature> getF_list() {
        return f_list;
    }

    public void setF_list(List<Feature> f_list) {
        this.f_list = f_list;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<Characteristic> getList() {
        return list;
    }

    public void setList(List<Characteristic> list) {
        this.list = list;
    }
}
