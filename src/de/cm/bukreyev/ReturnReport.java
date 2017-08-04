package de.cm.bukreyev;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PropertyTemplate;

import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Created by bukreyev on 02.08.2017.
 */
public class ReturnReport {
    private static int actID;
    private String data_time;
    private String just_data;
    private String externalOrder, boxId, carrier, senderID,
            boxConditionBase, boxConditionDetail, productData,
            productCategory, productCondition;

    public static void main(String[] args) {
        ReturnReport report = new ReturnReport();
        report.setCurrentDataTimeToString();
        report.generateFile();
    }

    private void setCurrentDataTimeToString() {
        Date dataTime = new Date();
        DateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
        data_time = dateFormat.format(dataTime);

        DateFormat dateFormat1 = new SimpleDateFormat("yyyy/MM/dd");
        just_data = dateFormat1.format(dataTime);
    }

    ReturnReport() {
        generateFile();
    }

    private void generateFile() {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("New sheet");

        fillRows(wb, sheet);
        borderCells(sheet);

        try {
            FileOutputStream fileOut = new FileOutputStream("C:\\Users\\bukreyev\\Desktop\\xl\\worklalala.xls");
            wb.write(fileOut);
            fileOut.close();
        } catch (java.io.IOException e) {
            e.printStackTrace();
        }
    }

    private void borderCells(Sheet sheet) {
        PropertyTemplate pt = new PropertyTemplate();
        pt.drawBorders(new CellRangeAddress(5, 10, 0, 1), BorderStyle.THIN, BorderExtent.INSIDE);
        pt.drawBorders(new CellRangeAddress(5, 10, 0, 1), BorderStyle.THIN, BorderExtent.OUTSIDE);

        pt.drawBorders(new CellRangeAddress(14, 16, 0, 1), BorderStyle.THIN, BorderExtent.INSIDE);
        pt.drawBorders(new CellRangeAddress(14, 16, 0, 1), BorderStyle.THIN, BorderExtent.OUTSIDE);

        pt.drawBorders(new CellRangeAddress(44, 44, 1, 1), BorderStyle.THIN, BorderExtent.BOTTOM);
        pt.drawBorders(new CellRangeAddress(47, 47, 1, 1), BorderStyle.THIN, BorderExtent.BOTTOM);
        pt.drawBorders(new CellRangeAddress(50, 50, 1, 1), BorderStyle.THIN, BorderExtent.BOTTOM);
        pt.applyBorders(sheet);
    }

    private void fillRows(Workbook wb, Sheet sheet) {
        Row row = sheet.createRow(0);
        Row row1 = sheet.createRow(3);
        Row row2 = sheet.createRow(5);
        Row row3 = sheet.createRow(6);
        Row row4 = sheet.createRow(7);
        Row row5 = sheet.createRow(8);
        Row row6 = sheet.createRow(9);
        Row row7 = sheet.createRow(10);
        Row row8 = sheet.createRow(12);
        Row row9 = sheet.createRow(14);
        Row row10 = sheet.createRow(15);
        Row row11 = sheet.createRow(16);

        Row row12 = sheet.createRow(44);
        Row row13 = sheet.createRow(47);
        Row row14 = sheet.createRow(50);

        Row row15 = sheet.createRow(45);
        Row row16 = sheet.createRow(48);
        Row row17 = sheet.createRow(51);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 1));
        sheet.addMergedRegion(new CellRangeAddress(12, 12, 0, 1));

        Cell cell0 = row.createCell(0);
        Cell cell_1 = row1.createCell(1);
        Cell cell2 = row8.createCell(0);
        Cell cell_3 = row15.createCell(1);

        CellStyle cellStyleAligCentr = wb.createCellStyle();
        cellStyleAligCentr.setAlignment(HorizontalAlignment.CENTER);

        CellStyle cellStyleAligRigh = wb.createCellStyle();
        cellStyleAligRigh.setAlignment(HorizontalAlignment.RIGHT);

        Font font = wb.createFont();
        font.setFontHeight((short) 6);

        CellStyle cellStyleSmallFont = wb.createCellStyle();
        cellStyleSmallFont.setFont(font);

        cell0.setCellStyle(cellStyleAligCentr);
        cell0.setCellValue(String.format("Акт осмотра № AI-%06d", actID++));

        cell2.setCellStyle(cellStyleAligCentr);
        cell2.setCellValue("Детальный контроль");

        row1.createCell(0).setCellValue(String.format("%s", "г. Москва"));
        cell_1.setCellStyle(cellStyleAligRigh);
        cell_1.setCellValue(String.format("%s", data_time));

        row2.createCell(0).setCellValue("Заказ №");
        externalOrder = "testdatatestdatatestdatatestdata";
        row2.createCell(1).setCellValue(externalOrder);
        row3.createCell(0).setCellValue("Номер короба");
        row3.createCell(1).setCellValue(boxId);
        row4.createCell(0).setCellValue("Перевозчик");
        row4.createCell(1).setCellValue(carrier);
        row5.createCell(0).setCellValue("Номер отправления");
        row5.createCell(1).setCellValue(senderID);
        row6.createCell(0).setCellValue("Общее состояние упаковки");
        row6.createCell(1).setCellValue(boxConditionBase);
        row7.createCell(0).setCellValue("Детальное состояние упаковки");
        row7.createCell(1).setCellValue(boxConditionDetail);
        row9.createCell(0).setCellValue("Информация о товаре");
        row9.createCell(1).setCellValue(productData);
        row10.createCell(0).setCellValue("Категория товара");
        row10.createCell(1).setCellValue(productCategory);
        row11.createCell(0).setCellValue("Детальное состояние товара");
        row11.createCell(1).setCellValue(productCondition);

        row12.createCell(0).setCellValue("Представитель получателя      М.П");
        row12.createCell(1).setCellValue("                          /                                   /            " + just_data);
        row13.createCell(0).setCellValue("Представитель компании");
        row13.createCell(1).setCellValue("                          /                                   /            " + just_data);
        row14.createCell(0).setCellValue("Охрана");
        row14.createCell(1).setCellValue("                          /                                   /            " + just_data);

        row15.createCell(1).setCellValue("        Подпись                      ФИО                           Дата");
        cell_3.setCellStyle(cellStyleSmallFont);
        row16.createCell(1).setCellValue("        Подпись                      ФИО                           Дата");
        row17.createCell(1).setCellValue("        Подпись                      ФИО                           Дата");

        Footer footer = sheet.getFooter();
        footer.setRight("Page "+ org.apache.poi.hssf.usermodel.HeaderFooter.page() + " of "+ HeaderFooter.numPages());

        sheet.autoSizeColumn(0);
        sheet.setColumnWidth(1, 14000);
    }
}
