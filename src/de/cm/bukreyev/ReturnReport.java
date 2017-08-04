package de.cm.bukreyev;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
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
    Date dataTime;

    private static int actID;
    private String data_time;
    private String just_data;
    private String externalOrder, carrier, senderID, productData, productCategory, productCondition;
    private String storeNumber = "23";
    private String storeAdress = "Г. Москва, улица пушкина дом Трампарампампам 123123 помещение кабинет";

    public static void main(String[] args) {
        ReturnReport report = new ReturnReport();
        report.setCurrentDataTimeToString();
        report.generateFile();
    }

    private void setCurrentDataTimeToString() {
        dataTime = new Date();
        DateFormat longFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
        data_time = longFormat.format(dataTime);

        DateFormat shortFormat = new SimpleDateFormat("yyyy/MM/dd");
        just_data = shortFormat.format(dataTime);
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
        pt.drawBorders(new CellRangeAddress(11, 13, 0, 1), BorderStyle.THIN, BorderExtent.INSIDE);
        pt.drawBorders(new CellRangeAddress(11, 13, 0, 1), BorderStyle.THIN, BorderExtent.OUTSIDE);

        pt.drawBorders(new CellRangeAddress(18, 20, 0, 1), BorderStyle.THIN, BorderExtent.INSIDE);
        pt.drawBorders(new CellRangeAddress(18, 20, 0, 1), BorderStyle.THIN, BorderExtent.OUTSIDE);

        pt.drawBorders(new CellRangeAddress(40, 40, 1, 1), BorderStyle.THIN, BorderExtent.BOTTOM);
        pt.drawBorders(new CellRangeAddress(43, 43, 1, 1), BorderStyle.THIN, BorderExtent.BOTTOM);
        pt.drawBorders(new CellRangeAddress(46, 46, 1, 1), BorderStyle.THIN, BorderExtent.BOTTOM);
        pt.applyBorders(sheet);
    }

    private void fillRows(Workbook wb, Sheet sheet) {
        Row rowAct = sheet.createRow(7);
        Row rowDataTime = sheet.createRow(1);
        Row rowStoreAdr = sheet.createRow(1);
        Row rowOrdId = sheet.createRow(11);
        Row rowCarrier = sheet.createRow(12);
        Row rowWayBill = sheet.createRow(13);
        Row rowDetailCntr = sheet.createRow(16);
        Row rowProdInfo = sheet.createRow(18);
        Row rowProdCtgr = sheet.createRow(19);
        Row rowProdCondtn = sheet.createRow(20);

        Row rowReceiver = sheet.createRow(40);
        Row rowCompMan = sheet.createRow(43);
        Row rowSecurMan = sheet.createRow(46);

        Row rowSign = sheet.createRow(41);
        Row rowSign1 = sheet.createRow(44);
        Row rowSign2 = sheet.createRow(47);

        sheet.addMergedRegion(new CellRangeAddress(7, 7, 0, 1));
        sheet.addMergedRegion(new CellRangeAddress(16, 16, 0, 1));

        Cell cellAct = rowAct.createCell(0);
        Cell cellDetailCntr = rowDetailCntr.createCell(0);
        Cell cellDataTime = rowDataTime.createCell(1);
        Cell cellStoreAdr = rowStoreAdr.createCell(0);


        CellStyle cellStyleAligCentr = wb.createCellStyle();
        cellStyleAligCentr.setAlignment(HorizontalAlignment.CENTER);

        CellStyle cellStyleAligRigh = wb.createCellStyle();
        cellStyleAligRigh.setAlignment(HorizontalAlignment.RIGHT);

        CellStyle cellStyleWrap = wb.createCellStyle();
        cellStyleWrap.setWrapText(true);

        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 6);

        CellStyle cellStyleSmallFont = wb.createCellStyle();
        cellStyleSmallFont.setFont(font);

        cellAct.setCellStyle(cellStyleAligCentr);
        cellAct.setCellValue(String.format("Акт осмотра № AI-%06d", actID++));

        cellDetailCntr.setCellStyle(cellStyleAligCentr);
        cellDetailCntr.setCellValue("Детальный контроль");

        cellStoreAdr.setCellValue(String.format
                ("Филиал \"Медиа Маркт № %s\"\nООО \"Медиа Маркт Сатурн\n\"%s", storeNumber, storeAdress));
        cellStoreAdr.setCellStyle(cellStyleWrap);

        cellDataTime.setCellValue(String.format("%s", data_time));
        cellDataTime.setCellStyle(cellStyleAligRigh);

        rowOrdId.createCell(0).setCellValue("Заказ №");
        rowOrdId.createCell(1).setCellValue(externalOrder);
        rowCarrier.createCell(0).setCellValue("Перевозчик");
        rowCarrier.createCell(1).setCellValue(carrier);
        rowWayBill.createCell(0).setCellValue("Номер отправления");
        rowWayBill.createCell(1).setCellValue(senderID);
        rowProdInfo.createCell(0).setCellValue("Информация о товаре");
        rowProdInfo.createCell(1).setCellValue(productData);
        rowProdCtgr.createCell(0).setCellValue("Категория товара");
        rowProdCtgr.createCell(1).setCellValue(productCategory);
        rowProdCondtn.createCell(0).setCellValue("Состояние товара");
        rowProdCondtn.createCell(1).setCellValue(productCondition);

        rowReceiver.createCell(0).setCellValue("Представитель получателя      М.П");
        rowReceiver.createCell(1).setCellValue("                          /                                   /            " + just_data);
        rowCompMan.createCell(0).setCellValue("Представитель компании");
        rowCompMan.createCell(1).setCellValue("                          /                                   /            " + just_data);
        rowSecurMan.createCell(0).setCellValue("Охрана ");
        rowSecurMan.createCell(1).setCellValue("                          /                                   /            " + just_data);

        rowSign.createCell(1).setCellValue("        Подпись                      ФИО                           Дата");
        rowSign1.createCell(1).setCellValue("        Подпись                      ФИО                           Дата");
        rowSign2.createCell(1).setCellValue("        Подпись                      ФИО                           Дата");
//        cell_3.setCellStyle(cellStyleSmallFont);

        Footer footer = sheet.getFooter();
        footer.setRight("Page " + org.apache.poi.hssf.usermodel.HeaderFooter.page() + " of " + HeaderFooter.numPages());

        sheet.setColumnWidth(0, 8000);
        sheet.setColumnWidth(1, 14000);
    }
}
