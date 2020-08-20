package com.zzqa.docx4j2word;

import com.zzqa.utils.Docx4jUtil;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.*;

import java.io.File;
import java.math.BigInteger;
import java.util.List;

/**
 * ClassName: AddingFooter
 * Description:
 *
 * @author 张文豪
 * @date 2020/8/3 10:38
 */
public class AddingAFooter {

    private ObjectFactory factory= Context.getWmlObjectFactory();


    public Relationship createFooterPart(WordprocessingMLPackage wordMLPackage ,String content) throws Exception {
        FooterPart footerPart = new FooterPart();

        footerPart.setPackage(wordMLPackage);

        P p = newImage(wordMLPackage, footerPart, "src\\main\\resources\\images\\中自庆安.png");

        //插入内容和图片
        //对字符串进行处理，如果出现&&标识则表示换行
        String[] split = {};
        if (content!=null){
            split = content.split("&&");
        }
        footerPart.setJaxbElement(createFooter(split,p));


        return wordMLPackage.getMainDocumentPart().addTargetPart(footerPart);
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
    private P newImage(WordprocessingMLPackage word,
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


    private Ftr createFooter(String[] split,P p) {
        Ftr footer = factory.createFtr();
        Tbl tbl = factory.createTbl();
        Tr tr = factory.createTr();
        Tc tc1 = factory.createTc();
        Tc tc2 = factory.createTc();
        //文本内容
        P paragraph = factory.createP();
        R run = factory.createR();
        RPr rPr = factory.createRPr();
        Docx4jUtil.setFontSize(rPr,"16");
        //去掉段落空行
        PPr paragraphPPr = paragraph.getPPr();
        if (paragraphPPr==null){
            paragraphPPr =factory.createPPr();
        }
        Docx4jUtil.setSpacing(paragraphPPr);
        run.getContent().add(rPr);

        for (int i=0;i<split.length;i++){
            Text text = new Text();
            text.setValue(split[i]);
            run.getContent().add(text);
            if (i!=split.length-1){
                Br br = new Br();
                run.getContent().add(br);
            }

        }

        paragraph.setPPr(paragraphPPr);
        paragraph.getContent().add(run);
        //设置插入图片对齐方式
        PPr pPr = p.getPPr();
        if (pPr==null){
            pPr = new PPr();
        }
        Jc jc = new Jc();
        jc.setVal(JcEnumeration.RIGHT);
        pPr.setJc(jc);
        p.setPPr(pPr);

        setCellWidth(tc1, 7500);
        setCellWidth(tc2, 2000);

        tc1.getContent().add(paragraph);
        tc2.getContent().add(p);

        tr.getContent().add(tc1);
        tr.getContent().add(tc2);
        tbl.getContent().add(tr);
//        footer.getContent().add(paragraph);
//        footer.getContent().add(p);

        footer.getContent().add(tbl);
        return footer;
    }


    public void createFooterReference(WordprocessingMLPackage wordMLPackage,Relationship relationship) {
        List<SectionWrapper> sections =
                wordMLPackage.getDocumentModel().getSections();

        SectPr sectionProperties = sections.get(sections.size() - 1).getSectPr();
        // There is always a section wrapper, but it might not contain a sectPr
        if (sectionProperties==null ) {
            sectionProperties = factory.createSectPr();
            wordMLPackage.getMainDocumentPart().addObject(sectionProperties);
            sections.get(sections.size() - 1).setSectPr(sectionProperties);
        }

        FooterReference footerReference = factory.createFooterReference();
        footerReference.setId(relationship.getId());
        footerReference.setType(HdrFtrRef.DEFAULT);
        sectionProperties.getEGHdrFtrReferences().add(footerReference);
    }

    /**
     *  本方法创建一个单元格属性集对象和一个表格宽度对象. 将给定的宽度设置到宽度对象然后将其添加到
     *  属性集对象. 最后将属性集对象设置到单元格中.
     */
    private void setCellWidth(Tc tableCell, int width) {
        TcPr tableCellProperties = new TcPr();
        TblWidth tableWidth = new TblWidth();
        tableWidth.setW(BigInteger.valueOf(width));
        tableCellProperties.setTcW(tableWidth);
        tableCell.setTcPr(tableCellProperties);
    }
}
