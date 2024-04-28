package com.newlight77.kata.survey.service;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileOutputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

@Service
public class FileService {
    private static final DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");

    public File saveWorkbookToFile(Workbook workbook, String filename) {
        File resultFile = new File(System.getProperty("java.io.tmpdir"), filename + "-" + dateTimeFormatter.format(LocalDate.now()) + ".xlsx");
        try (FileOutputStream outputStream = new FileOutputStream(resultFile)) {
            workbook.write(outputStream);
        } catch (Exception e) {
            throw new RuntimeException("Failed to write Excel file", e);
        }
        return resultFile;
    }
}
