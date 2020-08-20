package com.zzqa.docx4j2word;

import com.zzqa.utils.Docx4jUtil;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.wml.*;

import java.io.File;
import java.math.BigInteger;
import java.util.Date;

/**
 * ClassName: Test2
 * Description:
 *
 * @author 张文豪
 * @date 2020/7/31 14:30
 */
public class Cover {
    private ObjectFactory objectFactory = new ObjectFactory();

    public WordprocessingMLPackage createCover(WordprocessingMLPackage wpMLPackage,
                                               String reportName,Long startTime,Long endTime,
                                               String logoPath,String linePath){
        try {
            wpMLPackage = WordprocessingMLPackage.createPackage();

            //logo图片插入
            File logo = new File(logoPath);
            byte[] logoBytes = Docx4jUtil.convertImageToByteArray(logo);
            Docx4jUtil.addImageToPackage(wpMLPackage, logoBytes);

            //横线插入
            File linePng = new File(linePath);
            byte[] linePngBytes = Docx4jUtil.convertImageToByteArray(linePng);
            Docx4jUtil.addImageToPackage(wpMLPackage, linePngBytes);

            //加入换行
            Docx4jUtil.addBr(wpMLPackage,1);

            //封面标题
            addTitleWordToPackage(wpMLPackage, reportName);
            addTitleWordToPackage(wpMLPackage, Docx4jUtil.getDate(new Date(startTime))+"-"+Docx4jUtil.getDate(new Date(endTime)));
            addTitleWordToPackage(wpMLPackage, "检测报告");

            //插入换行符
            Docx4jUtil.addBr(wpMLPackage,2);
            //横线插入
            Docx4jUtil.addImageToPackage(wpMLPackage, linePngBytes);
            //插入换行符
            Docx4jUtil.addBr(wpMLPackage,1);

            //向封面插入小标题
            addTitleLittle(wpMLPackage,"浙江中自庆安新能源技术有限公司");
            addTitleLittle(wpMLPackage,"检测与诊断中心");
            addTitleLittle(wpMLPackage,"联系电话：4000093668-7");
            addTitleLittle(wpMLPackage,"邮箱：service@windit.com.cn");
            addTitleLittle(wpMLPackage, Docx4jUtil.getDate(new Date()));

            //下一页
            Docx4jUtil.addNextSection(wpMLPackage);

//            wpMLPackage.save(new File("D:/TestFile/文件1.docx"));
            //TODO 删除输出语句
            System.out.println("Cover Success......");

        } catch (Exception e) {
            e.printStackTrace();
        }
        return wpMLPackage;
    }

    /**
     * 创建包含图片的内容
     *
     * @param word
     * @param sourcePart
     * @param imageFilePath
     * @return
     * @throws Exception
     */
    public static P newImage(WordprocessingMLPackage word,
                             Part sourcePart,
                             String imageFilePath) throws Exception {
        BinaryPartAbstractImage imagePart = BinaryPartAbstractImage
                .createImagePart(word, sourcePart, new File(imageFilePath));
        //随机数ID
        int id = (int) (Math.random() * 10000);
        //这里的id不重复即可
        Inline inline = imagePart.createImageInline("image", "image", id, id * 2, false);

        ObjectFactory factory = new ObjectFactory();
        Drawing drawing = factory.createDrawing();
        drawing.getAnchorOrInline().add(inline);

        R r = factory.createR();
        r.getContent().add(drawing);

        P p = factory.createP();
        p.getContent().add(r);

        return p;
    }

    /**
     * 插入下一页
     * @param wpMLPackage
     */
    public void addNextPage(WordprocessingMLPackage wpMLPackage){
        P para = objectFactory.createP();
        SectPr sectPr = objectFactory.createSectPr();
        PPr pPr = objectFactory.createPPr();
        SectPr.Type sectPrType = objectFactory.createSectPrType();

        sectPrType.setVal("nextPage");
        sectPr.setType(sectPrType);
        pPr.setSectPr(sectPr);
        para.setPPr(pPr);
        wpMLPackage.getMainDocumentPart().addObject(para);
    }

    /**
     * 插入封面小标题
     * @param wpMLPackage
     * @param word
     */
    public void addTitleLittle(WordprocessingMLPackage wpMLPackage,String word){
        P para = objectFactory.createP();
        R run = objectFactory.createR();
        Text text = objectFactory.createText();
        //添加标题内容
        text.setValue(word);
        //设置居中格式
        PPr pPr = objectFactory.createPPr();
        Jc jc = pPr.getJc();
        if (jc==null){
            jc = new Jc();
        }
        jc.setVal(JcEnumeration.CENTER);

        RPr rPr = objectFactory.createRPr();
        //设置颜色
        Docx4jUtil.setFontColor(rPr, true, "#0070c0");
//        Color color = new Color();
//        color.setVal("#0070c0");
//        rPr.setColor(color);
        //设置字体大小
        HpsMeasure fontSize = new HpsMeasure();
        fontSize.setVal(new BigInteger("28"));
        rPr.setSzCs(fontSize);
        rPr.setSz(fontSize);

        //设置加粗
//        BooleanDefaultTrue booleanDefaultTrue = objectFactory.createBooleanDefaultTrue();
//        booleanDefaultTrue.setVal(Boolean.TRUE);
//        rPr.setB(booleanDefaultTrue);

        //设置字体
        RFonts rFonts = new RFonts();
        rFonts.setEastAsia("宋体");
        rPr.setRFonts(rFonts);


        //将格式和内容添加进para
        pPr.setJc(jc);
        run.getContent().add(rPr);
        run.getContent().add(text);
        para.setPPr(pPr);
        para.getContent().add(run);

        wpMLPackage.getMainDocumentPart().addObject(para);

    }

    /**
     * 封面标题
     * @param wpMLPackage
     * @param word 标题内容
     */
    public void addTitleWordToPackage(WordprocessingMLPackage wpMLPackage, String word){
//        ObjectFactory objectFactory = new ObjectFactory();
        P para = objectFactory.createP();
        R run = objectFactory.createR();
        Text text = objectFactory.createText();
        text.setValue(word);
        PPr pPr = objectFactory.createPPr();
        Jc jc = pPr.getJc();
        if (jc == null){
            jc = new Jc();
        }
        //设置居中
        jc.setVal(JcEnumeration.CENTER);
        pPr.setJc(jc);
        //设置字体
        RPr rPr = new RPr();
        rPr.setB(Context.getWmlObjectFactory().createBooleanDefaultTrue());
        Docx4jUtil.setFontSize(rPr, "50");
        Docx4jUtil.setFont(rPr, "宋体");
        run.setRPr(rPr);


        //将内容添加
        run.getContent().add(text);
        para.getContent().add(run);
        para.setPPr(pPr);
        wpMLPackage.getMainDocumentPart().addObject(para);

    }
}
