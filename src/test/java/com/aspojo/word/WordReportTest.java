package com.aspojo.word;


import cn.hutool.core.collection.CollectionUtil;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.junit.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;

import java.io.FileOutputStream;
import java.math.BigInteger;
import java.util.HashMap;
import java.util.Map;

public class WordReportTest {
    @Test
    public void testMergeCell() throws Exception {

        XWPFDocument document = new XWPFDocument();

        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText("The table:");

        // create table
        XWPFTable table = document.createTable(3, 5);

        for (int row = 0; row < 3; row++) {
            for (int col = 0; col < 5; col++) {
                table.getRow(row).getCell(col).setText("row " + row + ", col " + col);
            }
        }

        // create and set column widths for all columns in all rows
        // most examples don't set the type of the CTTblWidth but this
        // is necessary for working in all office versions
        for (int col = 0; col < 5; col++) {
            CTTblWidth tblWidth = CTTblWidth.Factory.newInstance();
            tblWidth.setW(BigInteger.valueOf(1000));
            tblWidth.setType(STTblWidth.DXA);
            for (int row = 0; row < 3; row++) {
                CTTcPr tcPr = table.getRow(row).getCell(col).getCTTc().getTcPr();
                if (tcPr != null) {
                    tcPr.setTcW(tblWidth);
                } else {
                    tcPr = CTTcPr.Factory.newInstance();
                    tcPr.setTcW(tblWidth);
                    table.getRow(row).getCell(col).getCTTc().setTcPr(tcPr);
                }
            }
        }

        // using the merge methods
        WordXParser.mergeCellVertically(table, 0, 0, 1);
        WordXParser.mergeCellVertically(table, 0, 1, 2);
        WordXParser.mergeCellHorizontally(table, 1, 2, 3);
        WordXParser.mergeCellHorizontally(table, 2, 1, 4);

        FileOutputStream out = new FileOutputStream("create_table.docx");
        document.write(out);

        System.out.println("create_table.docx written successully");
    }

    /***
     * ??????????????????????????????
     *
     */
    @Test
    public void testExportReport() {


        Map<String, Object> context = new HashMap<>();
        Map<String, Object> relation = new HashMap<>();
        context.put("list", CollectionUtil.toList(relation));


        relation.put("upstreamDocName", "????????????");
        relation.put("docName", "????????????");
        relation.put("upstreamDocTitle", "????????????");
        relation.put("docTitle", "????????????");
        relation.put("upReqTitle", "??????");
        relation.put("upSectionTitle", "??????");
        relation.put("reqTitle", "??????");
        relation.put("sectionTitle", "??????");

        Map<String, Object> req = new HashMap<>();
        relation.put("reqList", CollectionUtil.toList(req, req));

        req.put("reqName", "??????1");
        req.put("section", "??????1");
        req.put("upReqName", "??????1");
        req.put("upSection", "??????1");

        WordReport.exportDocx("template.docx", context, "report.docx");
    }
}