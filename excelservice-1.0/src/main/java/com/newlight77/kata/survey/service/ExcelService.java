package com.newlight77.kata.survey.service;
import com.newlight77.kata.survey.model.AddressStatus;
import com.newlight77.kata.survey.model.Campaign;
import com.newlight77.kata.survey.model.Survey;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class ExcelService {

    public Workbook createSurveyResultsWorkbook(Campaign campaign, Survey survey) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Survey");
        createColumnWidths(sheet);

        createHeader(workbook, sheet);
        addClientDetails(workbook, sheet, survey);
        addSurveyResults(workbook, sheet, campaign);
        return workbook;
    }

    private void createColumnWidths(Sheet sheet) {
        sheet.setColumnWidth(0, 10500);
        for (int i = 1; i <= 18; i++) {
            sheet.setColumnWidth(i, 6000);
        }
    }

    private void createHeader(Workbook workbook, Sheet sheet) {
        Row header = sheet.createRow(0);
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFFont font = ((XSSFWorkbook) workbook).createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 14);
        font.setBold(true);
        headerStyle.setFont(font);

        Cell headerCell = header.createCell(0);
        headerCell.setCellValue("Survey");
        headerCell.setCellStyle(headerStyle);
    }

    private void addClientDetails(Workbook workbook, Sheet sheet, Survey survey) {
        CellStyle titleStyle = createTitleCellStyle(workbook);
        CellStyle textStyle = createTextStyle(workbook);

        // Client Name
        Row row = sheet.createRow(2);
        Cell cell = row.createCell(0);
        cell.setCellValue("Client");
        cell.setCellStyle(titleStyle);

        Row clientRow = sheet.createRow(3);
        Cell clientNameCell = clientRow.createCell(0);
        clientNameCell.setCellValue(survey.getClient());
        clientNameCell.setCellStyle(textStyle);

        // Client Address
        String clientAddress = String.format("%s %s %s %s",
                survey.getClientAddress().getStreetNumber(),
                survey.getClientAddress().getStreetName(),
                survey.getClientAddress().getPostalCode(),
                survey.getClientAddress().getCity());
        Row addressRow = sheet.createRow(4);
        Cell addressCell = addressRow.createCell(0);
        addressCell.setCellValue(clientAddress);
        addressCell.setCellStyle(textStyle);
    }

    private void addSurveyResults(Workbook workbook, Sheet sheet, Campaign campaign) {
        CellStyle textStyle = createTextStyle(workbook);
        int rowIndex = 9;
        for (AddressStatus addressStatus : campaign.getAddressStatuses()) {
            Row row = sheet.createRow(rowIndex++);
            addSurveyResultRow(workbook, row, addressStatus, textStyle);
        }
    }

    private void addSurveyResultRow(Workbook workbook, Row row, AddressStatus addressStatus, CellStyle textStyle) {
        createCell(row, 0, addressStatus.getAddress().getStreetNumber(), textStyle);
        createCell(row, 1, addressStatus.getAddress().getStreetName(), textStyle);
        createCell(row, 2, addressStatus.getAddress().getPostalCode(), textStyle);
        createCell(row, 3, addressStatus.getAddress().getCity(), textStyle);
        createCell(row, 4, addressStatus.getStatus().toString(), textStyle);
    }

    private CellStyle createTitleCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFFont font = ((XSSFWorkbook) workbook).createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 12);
        font.setUnderline(FontUnderline.SINGLE);
        style.setFont(font);
        return style;
    }

    private CellStyle createTextStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setWrapText(true);
        return style;
    }

    private void createCell(Row row, int column, String value, CellStyle style) {
        Cell cell = row.createCell(column);
        cell.setCellValue(value);
        cell.setCellStyle(style);
    }
}