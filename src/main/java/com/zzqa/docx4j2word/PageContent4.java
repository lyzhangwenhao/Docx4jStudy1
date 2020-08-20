package com.zzqa.docx4j2word;

import com.zzqa.utils.Docx4jUtil;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.*;

import java.math.BigInteger;

/**
 * ClassName: PageContent4
 * Description:
 *
 * @author 张文豪
 * @date 2020/8/6 14:56
 */
public class PageContent4 {
    private ObjectFactory factory = new ObjectFactory();

    /**
     * 固定的第四部分内容
     * @param wpMLPackage
     */
    public void createPageContent4(WordprocessingMLPackage wpMLPackage){
        Relationship relationship = null;
        try {

            //添加标题四：补充说明
            wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading1", "4 补充说明");
            addParagraph(wpMLPackage, "(1) 本报告涂改无效。");
            addParagraph(wpMLPackage, "(2) 未经本中心书面许可，部分复制、摘用或篡改本报告内容，引起法律纠纷，责任自负。");
            addParagraph(wpMLPackage, "(3) 本监测报告是基于对机组所安装的CMS系统的振动数据所获得的信息自动生成的，因此，本报告对机组状态所做分析仅供参考。对于设备状况的最终判断以及所需采取的维护措施，可参考诊断分析报告，最终由用户自行决定。");
            addParagraph(wpMLPackage, "(4) 对检测报告若有异议，请于收到报告之日起一个月内向本中心提出，逾期不再受理。");

            Docx4jUtil.addBr(wpMLPackage, 20);
            addParagraph(wpMLPackage, "地址：杭州经济技术开发区6号路260号中自科技园");
            addParagraph(wpMLPackage, "邮编：310018");
            addParagraph(wpMLPackage, "电话：0571-28995840");
            addParagraph(wpMLPackage, "传真：0571-28995841");
            //TODO 删除输出语句
            System.out.println("PageContent4 Success......");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    /**
     * 添加正文段落
     * @param wpMLPackage
     * @param content
     */
    private void addParagraph(WordprocessingMLPackage wpMLPackage,String content){
        P p = factory.createP();
        R r = factory.createR();
        PPr pPr = p.getPPr();
        if (pPr==null){
            pPr = factory.createPPr();
        }


        //设置行距1.5倍
        PPrBase.Spacing spacing = pPr.getSpacing();
        if (spacing==null){
            spacing = factory.createPPrBaseSpacing();
        }
        spacing.setLineRule(STLineSpacingRule.AUTO);
        spacing.setLine(BigInteger.valueOf(360));
        spacing.setAfter(new BigInteger("0"));
        //段前段后
//        spacing.setBeforeLines(BigInteger.valueOf(50));
//        spacing.setAfterLines(BigInteger.valueOf(50));

        Text text = factory.createText();
        text.setValue(content);
        r.getContent().add(text);

        pPr.setSpacing(spacing);
        p.getContent().add(r);
        p.setPPr(pPr);
        wpMLPackage.getMainDocumentPart().addObject(p);
    }
}
