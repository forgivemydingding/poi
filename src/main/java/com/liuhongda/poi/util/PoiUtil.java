package com.liuhongda.poi.util;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.math.BigInteger;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.LinkedList;
import java.util.List;

/**
 * @author by liu.hongda
 * @Description TODO
 * @Date 2019/7/24 14:37
 */

public class PoiUtil {
    public static void main(String[] args) throws Exception {
        PoiUtil poiUtil = new PoiUtil();

        //创建文本例子
        //poiUtil.createTitle();

        //创建表例子
        //poiUtil.createTable();

        Calendar a = Calendar.getInstance();
        int year = a.get(Calendar.YEAR);
        int month = (a.get(Calendar.MONTH)) + 1;
        int day = a.get(Calendar.DATE);
        String date = year + "年" + month + "月" + day + "日";
        System.out.println(date);

    }

//    public void exportWord(List<Presentation> presentations, LocalDateTime time, String path) throws Exception {
//
//        //创建一个文档实例
//        XWPFDocument xwpfDocument = new XWPFDocument();
//
//        //添加标题
//        XWPFParagraph title = xwpfDocument.createParagraph();
//
//        //段落居中
//        title.setAlignment(ParagraphAlignment.CENTER);
//
//        XWPFRun run = title.createRun();
//        run.setText("各设区市" + time.getMonthValue() + "月份新闻发布会统计表");
//        run.setColor("000000");
//        run.setFontSize(20);
//
//        //换行
//        XWPFParagraph brParagraph = xwpfDocument.createParagraph();
//        XWPFRun brRun = brParagraph.createRun();
//        brRun.setText("\n");
//
//        for (AreaEnum areaEnum : AreaEnum.values()) {
//            LinkedList<Presentation> oneArea = new LinkedList<>();
//            String name = areaEnum.getName();
//            presentations.stream().forEach(p -> {
//                if (p.getTerritory().equals(name)) {
//                    oneArea.add(p);
//                }
//            });
//            if (!oneArea.isEmpty()) {
//                exportTable(oneArea, xwpfDocument, time);
//            }
//        }
//
//        File file = new File(path);
//        if (file.exists()) {
//            file.delete();
//        }
//        //文件不存在时会自动创建
//        OutputStream os = new FileOutputStream(file);
//        //写入文件
//        xwpfDocument.write(os);
//        os.close();
//
//    }
//
//    /**
//     * 输出table
//     *
//     * @param tableList
//     * @param xwpfDocument
//     * @param time
//     */
//    private void exportTable(List<Presentation> tableList, XWPFDocument xwpfDocument, LocalDateTime time) {
//
//        //获取字段数量
//        int mapSize = 6;
//        //创建表
//        XWPFTable table = xwpfDocument.createTable(tableList.size() + 2, mapSize);
//
//        //表格属性
//        CTTblPr tablePr = table.getCTTbl().addNewTblPr();
//        CTTblWidth ctTblWidth = tablePr.addNewTblW();
//        ctTblWidth.setW(BigInteger.valueOf(8000));
//        tablePr.addNewJc().setVal(STJc.CENTER);
//
//        //获取所有行元素
//        List<XWPFTableRow> rows = table.getRows();
//        //首行合并为一格
//        mergeCellsHorizontal(table, 0, 0, mapSize - 1);
//
//        //给标题行赋值
//        XWPFTableRow xwpfTableRowTitle = rows.get(0);
//        xwpfTableRowTitle.setHeight(500);
//        List<XWPFTableCell> tableCells = xwpfTableRowTitle.getTableCells();
//        XWPFTableCell xwpfTableCellTitle = tableCells.get(0);
//        String territory = tableList.get(0).getTerritory();
//        centerCell(xwpfTableCellTitle, "截至" + DateTimeUtil.chineseDate(time) + "，" + territory + "市共举办新闻发布会" + tableList.size() + "场。", true);
//
//        //给字段赋值
//        XWPFTableRow xwpfTableRowCol = rows.get(1);
//        xwpfTableRowCol.setHeight(500);
//        List<XWPFTableCell> tableCellsCol = xwpfTableRowCol.getTableCells();
//        centerCell(tableCellsCol.get(0), "序号", true);
//        tableCellsCol.get(0).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(600));
//        centerCell(tableCellsCol.get(1), "发布时间", true);
//        tableCellsCol.get(1).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(2000));
//        centerCell(tableCellsCol.get(2), "主发布人", true);
//        tableCellsCol.get(2).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(1000));
//        centerCell(tableCellsCol.get(3), "主持人", true);
//        tableCellsCol.get(3).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(1000));
//        centerCell(tableCellsCol.get(4), "发布主题", true);
//        tableCellsCol.get(4).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(1800));
//        centerCell(tableCellsCol.get(5), "图文实录链接", true);
//
//        //当前行元素所有格子
//        List<XWPFTableCell> cells;
//        int k = 1;
//        int rowSize = rows.size();
//        //数据填入表格
//        for (int i = 2; i < rowSize; i++) {
//            Presentation presentation = tableList.get(i - 2);
//            rows.get(i).setHeight(500);
//            cells = rows.get(i).getTableCells();
//            centerCell(cells.get(0), String.valueOf(k++), false);
//            centerCell(cells.get(1), DateTimeUtil.chineseDate(presentation.getPublishTime().toLocalDateTime()), false);
//            centerCell(cells.get(2), presentation.getMainPublisher(), false);
//            centerCell(cells.get(3), presentation.getCompere(), false);
//            centerCell(cells.get(4), presentation.getTheme(), false);
//            centerCell(cells.get(5), presentation.getImgAddress(), false);
//        }
//
//        //换行
//        XWPFParagraph brParagraph = xwpfDocument.createParagraph();
//        XWPFRun brRun = brParagraph.createRun();
//        brRun.setText("\n");
//
//    }

    /**
     * 居中
     *
     * @param xwpfTableCell
     * @param param
     * @param setBold
     */
    private void centerCell(XWPFTableCell xwpfTableCell, String param, Boolean setBold) {
        XWPFParagraph xwpfParagraph = xwpfTableCell.addParagraph();
        XWPFRun pRun0 = xwpfParagraph.createRun();
        pRun0.setText(param);
        if (setBold == true) {
            pRun0.setBold(true);
        }
        //垂直居中
        xwpfTableCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        //水平居中
        xwpfParagraph.setAlignment(ParagraphAlignment.CENTER);
        xwpfTableCell.removeParagraph(0);

    }

    public void createTable() throws Exception {
        //创建一个文档实例
        XWPFDocument xwpfDocument = new XWPFDocument();
        //创建一个 5 X 5 的表格
        XWPFTable table = xwpfDocument.createTable(5, 6);

        mergeCellsHorizontal(table, 0, 0, 6);

        //所有行元素
        List<XWPFTableRow> rows = table.getRows();

        rows.get(0).isCantSplitRow();

        //表格属性
        CTTblPr tablePr = table.getCTTbl().addNewTblPr();

        CTTblWidth ctTblWidth = tablePr.addNewTblW();

        ctTblWidth.setW(BigInteger.valueOf(8000));

        //当前行元素
        XWPFTableRow row;
        //当前行元素所有格子
        List<XWPFTableCell> cells;
        //单个格子
        XWPFTableCell cell;

        int rowSize = rows.size();
        int cellSize;
        for (int i = 0; i < rowSize; i++) {
            row = rows.get(i);
            //新增单元格
            row.addNewTableCell();
            //设置行的高度
            row.setHeight(500);
            //行属性
            //CTTrPr rowPr = row.getCtRow().addNewTrPr();
            //这种方式是可以获取到新增的cell的。
            //List<CTTc> list = row.getCtRow().getTcList();
            cells = row.getTableCells();
            cellSize = cells.size();
            for (int j = 0; j < cellSize; j++) {
                cell = cells.get(j);
//                if ((i + j) % 2 == 0) {
//                    //设置单元格的颜色
//                    //红色
//                    cell.setColor("ff0000");
//                } else {
//                    //蓝色
//                    cell.setColor("0000ff");
//                }
                //单元格属性
                CTTcPr cellPr = cell.getCTTc().addNewTcPr();
                cellPr.addNewVAlign().setVal(STVerticalJc.CENTER);
                if (j == 3) {
                    //设置宽度
                    cellPr.addNewTcW().setW(BigInteger.valueOf(3000));
                }
                if (i != 0) {
                    cell.setText(i + ", " + j);
                }
            }
        }
        String path = "W:/table.docx";
        File file = new File(path);
        if (file.exists()) {
            file.delete();
        }
        //文件不存在时会自动创建
        OutputStream os = new FileOutputStream(file);
        //写入文件
        xwpfDocument.write(os);
        os.close();
    }

    public void createTitle() throws IOException {
        //创建document对象
        XWPFDocument document = new XWPFDocument();

        //添加标题
        XWPFParagraph title = document.createParagraph();

        //段落居中
        title.setAlignment(ParagraphAlignment.CENTER);

        XWPFRun run = title.createRun();
        run.setText("这是一个标题");
        run.setColor("000000");
        run.setFontSize(20);

        //把doc输出到输出流
        OutputStream os = new FileOutputStream("W:/simpleWrite.docx");
        document.write(os);
        os.close();
    }


    public void exportSTToWord(String titleName, String[][] list, HttpServletResponse response) throws IOException {
        //创建document对象
        XWPFDocument document = new XWPFDocument();

        //添加标题
        XWPFParagraph titleParagraph = document.createParagraph();
        //设置段落居中
        titleParagraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun titleParagraphRun = titleParagraph.createRun();
        titleParagraphRun.setText(titleName);
        titleParagraphRun.setColor("000000");
        titleParagraphRun.setFontSize(20);
        //创建表格
        int row = list.length + 2;
        int column = list[0].length + 3;
        XWPFTable table = document.createTable(row, column);
        setTableWidth(table, "10000");

        //处理表达合并和数据填充

        //合并列
        mergeCellsVertically(table, 0, 0, 1);
        mergeCellsVertically(table, 2, 0, 1);
        mergeCellsVertically(table, column - 1, 0, 1);

        //合并行
        for (int i = 0; i < row - 2; i++) {
            mergeCellsHorizontal(table, i, 0, 1);
            mergeCellsHorizontal(table, i, 2, 4);
        }

        mergeCellsHorizontal(table, 0, 5, column - 2);
        mergeCellsHorizontal(table, row - 2, 0, 4);
        mergeCellsHorizontal(table, row - 1, 0, 4);
        mergeCellsHorizontal(table, row - 1, 5, column - 1);

        //填充数据
        XWPFTableRow rowIndex = table.getRow(0);
        XWPFTableCell cell = rowIndex.getCell(0);
        cell.setText(list[0][0]);
        XWPFTableCell cell1 = rowIndex.getCell(2);
        cell1.setText(list[0][1]);
        XWPFTableCell cell2 = rowIndex.getCell(5);
        cell2.setText("试题类型及题量");
        XWPFTableCell cell3 = rowIndex.getCell(column - 1);
        cell3.setText(list[0][list[0].length - 1]);

        XWPFTableRow rowIndex1 = table.getRow(1);
        for (int j = 5, i = 2; j < column - 1; j++, i++) {
            XWPFTableCell cell4 = rowIndex1.getCell(j);
            cell4.setText(list[0][i]);
        }

        for (int i = 2, ii = 1; i < row - 1; i++, ii++) {
            XWPFTableRow rowIndex2 = table.getRow(i);
            for (int j = 0, jj = 0; j < column - 1; j++) {
                XWPFTableCell cell4 = null;
                if (j == 0) {
                    cell4 = rowIndex2.getCell(j);
                    cell4.setText(list[ii][jj]);
                    jj++;
                }
                if (j == 2) {
                    cell4 = rowIndex2.getCell(j);
                    cell4.setText(list[ii][jj]);
                    jj++;
                } else if (j >= 5) {
                    cell4 = rowIndex2.getCell(j);
                    cell4.setText(list[ii][jj]);
                    jj++;
                }
            }
        }

        XWPFTableRow rowIndex3 = table.getRow(row - 1);
        XWPFTableCell cell5 = rowIndex3.getCell(0);
        cell5.setText("其他需要说明的问题");

        document.write(response.getOutputStream());
    }


    /***
     *  跨行合并
     * @param table
     * @param col  合并列
     * @param fromRow 起始行
     * @param toRow   终止行
     */
    private void mergeCellsVertically(XWPFTable table, int col, int fromRow, int toRow) {
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
            if (rowIndex == fromRow) {
                // The first merged cell is set with RESTART merge value
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
            }
        }
    }

    /***
     * 跨列合并
     * @param table
     * @param row 所合并的行
     * @param fromCell  起始列
     * @param toCell   终止列
     */
    private void mergeCellsHorizontal(XWPFTable table, int row, int fromCell, int toCell) {
        for (int cellIndex = fromCell; cellIndex <= toCell; cellIndex++) {
            XWPFTableCell cell = table.getRow(row).getCell(cellIndex);
            if (cellIndex == fromCell) {
                // The first merged cell is set with RESTART merge value
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
            }
        }
    }


    /***
     * 导出word 设置行宽
     * @param table
     * @param width
     */
    private void setTableWidth(XWPFTable table, String width) {
        CTTbl ttbl = table.getCTTbl();
        CTTblPr tblPr = ttbl.getTblPr() == null ? ttbl.addNewTblPr() : ttbl.getTblPr();
        CTTblWidth tblWidth = tblPr.isSetTblW() ? tblPr.getTblW() : tblPr.addNewTblW();
        CTJc cTJc = tblPr.addNewJc();
        cTJc.setVal(STJc.Enum.forString("center"));
        tblWidth.setW(new BigInteger(width));
        tblWidth.setType(STTblWidth.DXA);
    }

}
