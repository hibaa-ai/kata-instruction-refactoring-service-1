package com.newlight77.kata.survey.service;

import com.newlight77.kata.survey.client.CampaignClient;
import com.newlight77.kata.survey.model.Campaign;
import com.newlight77.kata.survey.model.Survey;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Component;

import java.io.File;

@Component
public class ExportCampaignService {

  private final CampaignClient campaignClient;
  private final MailService mailService;
  private final ExcelService excelService;
  private final FileService fileService;

  public ExportCampaignService(CampaignClient campaignClient, MailService mailService, ExcelService excelService, FileService fileService) {
      this.campaignClient = campaignClient;
      this.mailService = mailService;
      this.excelService = excelService;
      this.fileService = fileService;
  }
  public void createSurvey(Survey survey) {
    campaignClient.createSurvey(survey);
  }

  public Survey getSurvey(String id) {
    return campaignClient.getSurvey(id);
  }

  public void createCampaign(Campaign campaign) {
    campaignClient.createCampaign(campaign);
  }

  public Campaign getCampaign(String id) {
    return campaignClient.getCampaign(id);
  }
  public void sendResults(final Campaign campaign, final Survey survey) {
    Workbook workbook = excelService.createSurveyWorkbook(survey, campaign);
    File file = fileService.saveWorkbookToFile(workbook, "survey-" + survey.getId());
    mailService.sendEmailWithAttachment("Campaign Results", "Hi,\n\nYou will find in the attached file the campaign results.", file);
  }
}