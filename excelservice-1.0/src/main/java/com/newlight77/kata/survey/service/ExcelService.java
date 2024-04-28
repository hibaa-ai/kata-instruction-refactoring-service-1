package com.newlight77.kata.survey.service;

import com.newlight77.kata.survey.model.AddressStatus;
import com.newlight77.kata.survey.model.Campaign;
import com.newlight77.kata.survey.model.Survey;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

@Service
public class ExcelService {

    public Workbook createSurveyWorkbook(Survey survey, Campaign campaign) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Survey");

        createHeader(sheet);
        createClientSection(sheet, survey);
        createSurveyDetails(sheet, campaign);

        return workbook;
    }

    private void createHeader(Sheet sheet) {
        Row header = sheet.createRow(0);
        CellStyle headerStyle = sheet.getWorkbook().createCellStyle();
        XSSFFont font = ((XSSFWorkbook) sheet.getWorkbook()).createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 14);
        font.setBold(true);
        headerStyle.setFont(font);
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setWrapText(false);

        Cell headerCell = header.createCell(0);
        headerCell.setCellValue("Survey");
        headerCell.setCellStyle(headerStyle);
    }

    private void createClientSection(Sheet sheet, Survey survey) {
        Row row = sheet.createRow(2);
        Cell cell = row.createCell(0);
        cell.setCellValue("Client");
        cell.setCellStyle(titleStyle(sheet));

        Row clientRow = sheet.createRow(3);
        Cell clientCell = clientRow.createCell(0);
        clientCell.setCellValue(survey.getClient());
        clientCell.setCellStyle(defaultStyle(sheet));

        String clientAddress = formatAddress(survey);
        Row addressRow = sheet.createRow(4);
        Cell addressCell = addressRow.createCell(0);
        addressCell.setCellValue(clientAddress);
        addressCell.setCellStyle(defaultStyle(sheet));
    }

    private String formatAddress(Survey survey) {
        return survey.getClientAddress().getStreetNumber() + " "
                + survey.getClientAddress().getStreetName() + ", "
                + survey.getClientAddress().getPostalCode() + " "
                + survey.getClientAddress().getCity();
    }

    private void createSurveyDetails(Sheet sheet, Campaign campaign) {
        Row row = sheet.createRow(6);
        Cell cell = row.createCell(0);
        cell.setCellValue("Number of surveys");
        cell = row.createCell(1);
        cell.setCellValue(campaign.getAddressStatuses().size());

        int rowIndex = 9;
        for (AddressStatus addressStatus : campaign.getAddressStatuses()) {
            Row surveyRow = sheet.createRow(rowIndex++);
            populateAddressStatusRow(surveyRow, addressStatus);
        }
    }

    private void populateAddressStatusRow(Row row, AddressStatus addressStatus) {
        row.createCell(0).setCellValue(addressStatus.getAddress().getStreetNumber());
        row.createCell(1).setCellValue(addressStatus.getAddress().getStreetName());
        row.createCell(2).setCellValue(addressStatus.getAddress().getPostalCode());
        row.createCell(3).setCellValue(addressStatus.getAddress().getCity());
        row.createCell(4).setCellValue(addressStatus.getStatus().toString());
    }

    private CellStyle titleStyle(Sheet sheet) {
        CellStyle style = sheet.getWorkbook().createCellStyle();
        XSSFFont font = ((XSSFWorkbook) sheet.getWorkbook()).createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 12);
        font.setUnderline(FontUnderline.SINGLE);
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    private CellStyle defaultStyle(Sheet sheet) {
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setWrapText(true);
        return style;
    }
}
