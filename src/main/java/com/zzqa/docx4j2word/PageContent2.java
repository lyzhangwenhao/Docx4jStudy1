package com.zzqa.docx4j2word;

import com.zzqa.pojo.Characteristic;
import com.zzqa.pojo.Feature;
import com.zzqa.pojo.UnitInfo;
import com.zzqa.utils.Docx4jUtil;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

/**
 * ClassName: PageContent2
 * Description:
 *
 * @author 张文豪
 * @date 2020/8/5 13:55
 */
public class PageContent2 {
    private ObjectFactory factory = new ObjectFactory();

    /**
     * 生成预警/报警的表格信息
     *
     * @param wpMLPackage 传入的wpMLPackage对象
     * @param characteristicList   生成表格的数据，其中UnitInfo对象里面的四个colKeys与data[0]必须一一对应，并且按照colKeys升序排列(包括时间，字符串，数值)数据类型为String
     */
    public void createPageContent2(WordprocessingMLPackage wpMLPackage, List<Characteristic> characteristicList) {
        //添加标题一：项目概述
        wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading1", "2 运行状况");

        Docx4jUtil.addTableTitle(wpMLPackage, "表1 预/报警机组统计");
        //创建一个表格
        Tbl tbl = factory.createTbl();
        //生成表头
        createTalbeTitle(tbl);
        //生成数据
        createTableData(wpMLPackage, tbl, characteristicList);
        //TODO 删除输出语句
        System.out.println("PageContent2 Success......");
        //跨行合并
//        mergeCellsVertically(tbl,0, 1, 2);
    }


    /**
     * 生成表格数据
     *
     * @param wpMLPackage
     * @param tbl
     * @param characteristicList
     */
    private void createTableData(WordprocessingMLPackage wpMLPackage, Tbl tbl, List<Characteristic> characteristicList) {
        if (characteristicList != null) {
            Tr dateTr = null;

            int row1 = 1;
            int row2 = 1;
            int row3 = 1;
            for (Characteristic characteristic1 : characteristicList) {
                int currentLine = 0;
                //部件集合
                List<Characteristic> list1 = characteristic1.getList();
                //机组名称
                String name1 = characteristic1.getName();
                for (Characteristic characteristic2 : list1) {
                    //测点集合
                    List<Characteristic> list2 = characteristic2.getList();
                    //部件名称
                    String name2 = characteristic2.getName();
                    int num = 0;
                    for (Characteristic characteristic3 : list2) {
                        //特征值数据集合
                        List<Feature> f_list = characteristic3.getF_list();
                        //测点名称
                        String name3 = characteristic3.getName();
                        for (Feature feature:f_list){
                            if ("预警".equals(feature.getLevel()) || "报警".equals(feature.getLevel())) {
                                dateTr = factory.createTr();
                                addTableTc(dateTr, name1, 1100, false, "black");
                                addTableTc(dateTr, name2, 1500, false, "black");
                                addTableTc(dateTr, name3, 1500, false, "black");
                                addTableTc(dateTr, String.valueOf(feature.getCount()), 1100, false, "black");
                                addTableTc(dateTr, feature.getValve(), 1100, false, "black");
                                addTableTc(dateTr, feature.getCodeName(), 1100, false, "black");
                                addTableTc(dateTr, feature.getMaxValue(), 1800, false, "black");
                                //根据报警等级添加
                                if ("预警".equals(feature.getLevel())) {
                                    addTableTc(dateTr, feature.getLevel(), 1100, true, "#92d050");
                                } else if ("报警".equals(feature.getLevel())) {
                                    addTableTc(dateTr, feature.getLevel(), 1100, true, "#ff0000");
                                }
                                currentLine++;
                                num++;
                                //将tr添加到表格中
                                tbl.getContent().add(dateTr);

                            }
                        }
                        mergeCellsVertically(tbl, 2, row3, row3+f_list.size());
                        row3 += f_list.size();
                    }
                    mergeCellsVertically(tbl, 1, row2, row2+num);
                    row2 += num;
                }
                mergeCellsVertically(tbl, 0, row1, row1+currentLine);
                row1 += currentLine;
            }
        }
        wpMLPackage.getMainDocumentPart().addObject(tbl);
    }

    /**
     * 为table生成表头
     *
     * @param tbl
     */
    private void createTalbeTitle(Tbl tbl) {
        //给table添加边框
        addBorders(tbl);
        Tr tr = factory.createTr();
        //单元格居中对齐
        Jc jc = new Jc();
        jc.setVal(JcEnumeration.CENTER);
        TblPr tblPr = tbl.getTblPr();
        tblPr.setJc(jc);
        tbl.setTblPr(tblPr);

        //表格表头
        addTableTc(tr, "报警机组", 1100, true, "black");
        addTableTc(tr, "报警部件", 1500, true, "black");
        addTableTc(tr, "报警测点", 1500, true, "black");
        addTableTc(tr, "报警次数", 1100, true, "black");
        addTableTc(tr, "报警门限", 1100, true, "black");
        addTableTc(tr, "报警类型", 1100, true, "black");
        addTableTc(tr, "最大报警特征值", 1800, true, "black");
        addTableTc(tr, "报警等级", 1100, true, "black");
        //将tr添加到table中
        tbl.getContent().add(tr);
    }

    /**
     * 添加TableCell
     *
     * @param tableRow
     * @param content
     */
    private void addTableTc2(Tr tableRow, String content, int width, boolean isBold, String color) {
        Tc tc = factory.createTc();
        P p = factory.createP();
        R r = factory.createR();
        RPr rPr = factory.createRPr();
        Text text = factory.createText();
        BooleanDefaultTrue bCs = rPr.getBCs();
        if (bCs == null) {
            bCs = new BooleanDefaultTrue();
        }
        bCs.setVal(true);
        rPr.setBCs(bCs);

        //设置宽度
        setCellWidth(tc, width);
        //生成段落添加到单元格中
        text.setValue(content);
        //设置字体颜色，加粗
        Docx4jUtil.setFontColor(rPr, isBold, color);
        //设置字体
        Docx4jUtil.setFont(rPr, "宋体");
        //设置字体大小
        Docx4jUtil.setFontSize(rPr, "20");
        //将样式添加到段落中
        r.getContent().add(rPr);

        r.getContent().add(text);
        p.getContent().add(r);
        tc.getContent().add(p);
        //设置垂直居中
        setTcVAlign(tc, STVerticalJc.CENTER);
        //设置水平居中
        setTcJcAlign(tc, JcEnumeration.CENTER);
        //去除段后格式
        PPr pPr = p.getPPr();
        if (pPr == null) {
            pPr = factory.createPPr();
        }
        Docx4jUtil.setSpacing(pPr);
        p.setPPr(pPr);
//        setCellVMerage(tc, currentRow, startRow, rowSpan, 2, 2);

        tableRow.getContent().add(tc);
    }

    /**
     * 添加TableCell
     *
     * @param tableRow
     * @param content
     */
    private void addTableTc(Tr tableRow, String content, int width, boolean isBold, String color) {
        Tc tc = factory.createTc();
        P p = factory.createP();
        R r = factory.createR();
        RPr rPr = factory.createRPr();
        Text text = factory.createText();
        BooleanDefaultTrue bCs = rPr.getBCs();
        if (bCs == null) {
            bCs = new BooleanDefaultTrue();
        }
        bCs.setVal(true);
        rPr.setBCs(bCs);

        //设置宽度
        setCellWidth(tc, width);
        //生成段落添加到单元格中
        text.setValue(content);
        //设置字体颜色，加粗
        Docx4jUtil.setFontColor(rPr, isBold, color);
        //设置字体
        Docx4jUtil.setFont(rPr, "宋体");
        //设置字体大小
        Docx4jUtil.setFontSize(rPr, "20");
        //将样式添加到段落中
        r.getContent().add(rPr);

        r.getContent().add(text);
        p.getContent().add(r);
        tc.getContent().add(p);
        //设置垂直居中
        setTcVAlign(tc, STVerticalJc.CENTER);
        //设置水平居中
        setTcJcAlign(tc, JcEnumeration.CENTER);
        //去除段后格式
        PPr pPr = p.getPPr();
        if (pPr == null) {
            pPr = factory.createPPr();
        }
        Docx4jUtil.setSpacing(pPr);
        p.setPPr(pPr);

        tableRow.getContent().add(tc);
    }

    /**
     * 本方法创建一个单元格属性集对象和一个表格宽度对象. 将给定的宽度设置到宽度对象然后将其添加到
     * 属性集对象. 最后将属性集对象设置到单元格中.
     */
    private void setCellWidth(Tc tableCell, int width) {
        TcPr tableCellProperties = new TcPr();
        TblWidth tableWidth = new TblWidth();
        tableWidth.setW(BigInteger.valueOf(width));
        tableCellProperties.setTcW(tableWidth);
        tableCell.setTcPr(tableCellProperties);
    }

    /**
     * 本方法为表格添加边框
     *
     * @param table
     */
    private void addBorders(Tbl table) {
        table.setTblPr(new TblPr());
        CTBorder border = new CTBorder();
        border.setColor("#95b3d7");
        border.setSz(new BigInteger("4"));
        border.setSpace(new BigInteger("0"));
        border.setVal(STBorder.SINGLE);

        TblBorders borders = new TblBorders();
        borders.setBottom(border);
        borders.setLeft(border);
        borders.setRight(border);
        borders.setTop(border);
        borders.setInsideH(border);
        borders.setInsideV(border);
        table.getTblPr().setTblBorders(borders);
    }

    private TcPr getTcPr(Tc tc) {
        TcPr tcPr = tc.getTcPr();
        if (tcPr == null) {
            tcPr = new TcPr();
            tc.setTcPr(tcPr);
        }
        return tcPr;
    }

    /**
     * @Description: 设置单元格垂直对齐方式
     */
    private void setTcVAlign(Tc tc, STVerticalJc vAlignType) {
        if (vAlignType != null) {
            TcPr tcPr = getTcPr(tc);
            CTVerticalJc vAlign = new CTVerticalJc();
            vAlign.setVal(vAlignType);
            tcPr.setVAlign(vAlign);
        }
    }

    /**
     * @Description: 设置单元格水平对齐方式
     */
    private void setTcJcAlign(Tc tc, JcEnumeration jcType) {
        if (jcType != null) {
            List<P> pList = getTcAllP(tc);
            for (P p : pList) {
                setParaJcAlign(p, jcType);
            }
        }
    }

    private List<P> getTcAllP(Tc tc) {
        List<Object> objList = getAllElementFromObject(tc, P.class);
        List<P> pList = new ArrayList<P>();
        if (objList == null) {
            return pList;
        }
        for (Object obj : objList) {
            if (obj instanceof P) {
                P p = (P) obj;
                pList.add(p);
            }
        }
        return pList;
    }

    /**
     * @Description: 得到指定类型的元素
     */
    private static List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
        List<Object> result = new ArrayList<Object>();
        if (obj instanceof JAXBElement)
            obj = ((JAXBElement<?>) obj).getValue();
        if (obj.getClass().equals(toSearch))
            result.add(obj);
        else if (obj instanceof ContentAccessor) {
            List<?> children = ((ContentAccessor) obj).getContent();
            for (Object child : children) {
                result.addAll(getAllElementFromObject(child, toSearch));
            }
        }
        return result;
    }

    /**
     * @Description: 设置段落水平对齐方式
     */
    private void setParaJcAlign(P paragraph, JcEnumeration hAlign) {
        if (hAlign != null) {
            PPr pprop = paragraph.getPPr();
            if (pprop == null) {
                pprop = new PPr();
                paragraph.setPPr(pprop);
            }
            Jc align = new Jc();
            align.setVal(hAlign);
            pprop.setJc(align);
        }
    }

    /**
     * @Description: 跨行合并
     */
    private void mergeCellsVertically(Tbl tbl, int col, int fromRow, int toRow) {
        if (col < 0 || fromRow < 0 || toRow < 0) {
            return;
        }
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            Tc tc = getTc(tbl, rowIndex, col);
            if (tc == null) {
                break;
            }
            TcPr tcPr = getTcPr(tc);
            TcPrInner.VMerge vMerge = tcPr.getVMerge();
            if (vMerge == null) {
                vMerge = new TcPrInner.VMerge();
                tcPr.setVMerge(vMerge);
            }
            if (rowIndex == fromRow) {
                vMerge.setVal("restart");
            } else {
                vMerge.setVal("continue");
            }
        }
    }

    /**
     * @Description: 得到表格所有的行
     */
    private List<Tr> getTblAllTr(Tbl tbl) {
        List<Object> objList = getAllElementFromObject(tbl, Tr.class);
        List<Tr> trList = new ArrayList<Tr>();
        if (objList == null) {
            return trList;
        }
        for (Object obj : objList) {
            if (obj instanceof Tr) {
                Tr tr = (Tr) obj;
                trList.add(tr);
            }
        }
        return trList;

    }

    /**
     * @Description:得到指定位置的单元格
     */
    private Tc getTc(Tbl tbl, int row, int cell) {
        if (row < 0 || cell < 0) {
            return null;
        }
        List<Tr> trList = getTblAllTr(tbl);
        if (row >= trList.size()) {
            return null;
        }
        List<Tc> tcList = getTrAllCell(trList.get(row));
        if (cell >= tcList.size()) {
            return null;
        }
        return tcList.get(cell);
    }

    /**
     * @Description:得到所有表格
     */
    private List<Tbl> getAllTbl(WordprocessingMLPackage wordMLPackage) {
        MainDocumentPart mainDocPart = wordMLPackage.getMainDocumentPart();
        List<Object> objList = getAllElementFromObject(mainDocPart, Tbl.class);
        if (objList == null) {
            return null;
        }
        List<Tbl> tblList = new ArrayList<Tbl>();
        for (Object obj : objList) {
            if (obj instanceof Tbl) {
                Tbl tbl = (Tbl) obj;
                tblList.add(tbl);
            }
        }
        return tblList;
    }

    /**
     * @Description: 获取所有的单元格
     */
    private List<Tc> getTrAllCell(Tr tr) {
        List<Object> objList = getAllElementFromObject(tr, Tc.class);
        List<Tc> tcList = new ArrayList<Tc>();
        if (objList == null) {
            return tcList;
        }
        for (Object tcObj : objList) {
            if (tcObj instanceof Tc) {
                Tc objTc = (Tc) tcObj;
                tcList.add(objTc);
            }
        }
        return tcList;
    }


}
