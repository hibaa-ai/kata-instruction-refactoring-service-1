package com.newlight77.kata.survey.service;

import com.newlight77.kata.survey.client.CampaignClient;
import com.newlight77.kata.survey.model.AddressStatus;
import com.newlight77.kata.survey.model.Campaign;
import com.newlight77.kata.survey.model.Survey;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileOutputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

@Service
public class ExportCampaignService {

  private final CampaignClient campaignWebService;
  private final MailService mailService;
  private final ExcelService excelService;
  private static final DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");

  public ExportCampaignService(CampaignClient campaignWebService, MailService mailService, ExcelService excelService) {
    this.campaignWebService = campaignWebService;
    this.mailService = mailService;
    this.excelService = excelService;
  }

  public void createSurvey(Survey survey) {
    campaignWebService.createSurvey(survey);
  }

  public Survey getSurvey(String id) {
    return campaignWebService.getSurvey(id);
  }

  public void createCampaign(Campaign campaign) {
    campaignWebService.createCampaign(campaign);
  }

  public Campaign getCampaign(String id) {
    return campaignWebService.getCampaign(id);
  }

  public void sendResults(Campaign campaign, Survey survey) {
    Workbook workbook = excelService.createSurveyResultsWorkbook(campaign, survey);
    writeFileAndSend(survey, workbook);
  }

  protected void writeFileAndSend(Survey survey, Workbook workbook) {
    try {
      File resultFile = new File(System.getProperty("java.io.tmpdir"), "survey-" + survey.getId() + "-" + dateTimeFormatter.format(LocalDate.now()) + ".xlsx");
      FileOutputStream outputStream = new FileOutputStream(resultFile);
      workbook.write(outputStream);
      outputStream.close();

      mailService.send(resultFile);
      resultFile.deleteOnExit();
    } catch (Exception ex) {
      throw new RuntimeException("Error while trying to send email", ex);
    } finally {
      try {
        workbook.close();
      } catch (Exception e) {
        // Log error if needed
      }
    }
  }
}