package org.testrail;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.microsoft.playwright.APIRequest;
import com.microsoft.playwright.APIRequestContext;
import com.microsoft.playwright.APIResponse;
import com.microsoft.playwright.Playwright;
import com.microsoft.playwright.options.RequestOptions;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.junit.Assert;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;


public class GetAttachmentsFromTestRail {

    public static int rowExcel = -1;
    public static List<String> allTestIDs = new ArrayList<>();
    public static List<String> allcaseIDs = new ArrayList<>();
    public static List<String> allcaseTitlesList = new ArrayList<>();
    public static List<String> allTestIDsInExcel = new ArrayList<>();
    public static final String pathForAttachmentsFolder = System.getProperty("user.dir") + "\\src\\main\\java\\org\\attachmentsfolder\\";
    public final String filePathforExcel = pathForAttachmentsFolder + "attachments.xlsx";
    public static final String filePathforAllJPGs = pathForAttachmentsFolder + "allJPGs\\";
    public final String BASE_URL = "https://nupsys.testrail.io/index.php?/api/v2/";
    public final String basic64Credentials = "basic bWt1cnVjYXlAbnVwc3lzLmNvbTpRdWFsaXR5MTIu";
    public final String RUN_ID = "117";
    @BeforeClass()
    public static void beforeClass() throws Exception {
        deleteAllJPGFiles(filePathforAllJPGs);
    }

    @Test()
    public void getAllAttachments() throws IOException {
        Playwright playwright = Playwright.create();

        APIRequestContext apiRequestContext = playwright.request().newContext(new APIRequest.NewContextOptions()
                .setIgnoreHTTPSErrors(true));

        // Get information for all tests in the RUN - please enter run id for the RUN above
        int index = -1;
        int numberofTests;
        while (true) {
            index++;
            String url = BASE_URL+"get_tests/"+RUN_ID+"&limit=1&offset="+index;
            List<String> id_caseid_list = null;
            try {
                id_caseid_list = getTestsForRun(apiRequestContext, url);

                System.out.println("id_caseid_title_list ["+index+"]= " + id_caseid_list);

                allTestIDs.add(id_caseid_list.get(0));
                allcaseIDs.add(id_caseid_list.get(1));
                allcaseTitlesList.add(id_caseid_list.get(2));
            } catch (Exception ignored) {
                numberofTests = index;
                System.out.println("All tests retrieved. Number of tests in the Run : " + numberofTests);
                break;
            }
        }


        // read all test IDs in the Excel file
        allTestIDsInExcel = readAllColumnFromExcel(filePathforExcel, 1);

        // for each test id in the list, get the latest attachment, downlaod it, and create hyperlink to the JPG file in the Excel
        for (int i = 0; i < allTestIDs.size(); i++) {
            Assert.assertTrue(allTestIDsInExcel.get(i).contains(allTestIDs.get(i)));

            String url = BASE_URL + "get_attachments_for_test/" +allTestIDs.get(i);
            try {
                String latestAttachmentID = getAttachmentForTest(apiRequestContext, url);

                url = BASE_URL + "get_attachment/" + latestAttachmentID;
                downloadAttachment(apiRequestContext, filePathforAllJPGs, url, latestAttachmentID);

                addHyperlinktoExcel(filePathforAllJPGs + latestAttachmentID + ".jpg");
            } catch (Exception e) {
                System.out.println("No attachment for this test ID");
                rowExcel++;
            }
        }
    }

    public static void deleteAllJPGFiles(String folderName) {
        File f = new File(folderName);

        String[] files;
        try {
            files = f.list();
            for (String file : files) {
                File currentFile = new File(f.getPath(), file);
                System.out.println("file deleted. " + currentFile.delete());
            }
        } catch (Exception ignored) {
            System.out.println("file not deleted");
        }
    }

    public String getAttachmentForTest(APIRequestContext apiRequestContext, String url) throws JsonProcessingException {
        APIResponse response = apiRequestContext.get(url, RequestOptions.create()
                .setHeader("Content-Type", "application/json")
                .setHeader("Authorization", basic64Credentials));

        return readFromJSONandGetLatestID(response, "id");
    }

    public static String readFromJSONandGetLatestID(APIResponse response, String key1) throws JsonProcessingException {
        String responseText = response.text();
        ObjectMapper objectMapper = new ObjectMapper();
        JsonNode rootnode = objectMapper.readTree(responseText);
        JsonNode attachmentsNode = rootnode.get(0);
        String latestAttachmentID = attachmentsNode.path(key1).asText();
        System.out.println("latestAttachmentID = " + latestAttachmentID);
        return latestAttachmentID;
    }


    public List<String> getTestsForRun(APIRequestContext apiRequestContext, String url) throws JsonProcessingException {
        APIResponse response = apiRequestContext.get(url, RequestOptions.create()
                .setHeader("Content-Type", "application/json")
                .setHeader("Authorization", basic64Credentials));

        List<String> idCaseIdandTitlesList = new ArrayList<>();
        if (response.ok()) {
            idCaseIdandTitlesList = readFromJSONandAddtoList(response, "tests", "id", "case_id" , "title");;
        }

        return idCaseIdandTitlesList;
    }

    public static List<String> readFromJSONandAddtoList(APIResponse response, String header, String key1, String key2, String key3) throws JsonProcessingException {
        String responseText = response.text();
        ObjectMapper objectMapper = new ObjectMapper();
        JsonNode rootnode = objectMapper.readTree(responseText);
        JsonNode testsNode = rootnode.path(header);
        JsonNode firstTest = testsNode.get(0);

        String id = firstTest.path(key1).asText();
        String case_id = firstTest.path(key2).asText();
        String title = firstTest.path(key3).asText();

        List<String> list = new ArrayList<>();
        list.add(id);
        list.add(case_id);
        list.add(title);

        return list;
    }

    public void downloadAttachment(APIRequestContext apiRequestContext, String folder, String url, String fileName) throws IOException {
        APIResponse response = apiRequestContext.get(url, RequestOptions.create()
                .setHeader("Content-Type", "application/json")
                .setHeader("Authorization", basic64Credentials));

        byte[] responseData = response.body();
        Files.write(Paths.get(folder + "\\" + fileName + ".jpg"), responseData);
    }

    public static List<String> readAllColumnFromExcel(String filePath, int cell) throws IOException {
        FileInputStream fis = null;
        List<String> valueOfString = new ArrayList<>();
        try {
            fis = new FileInputStream(filePath);
        } catch (FileNotFoundException ignored) {
            fis = new FileInputStream(filePath);
        }

        Workbook workbook = WorkbookFactory.create(fis);
        Sheet sheet = workbook.getSheetAt(0);
        int lastRowNumber = workbook.getSheetAt(0).getLastRowNum();

        for (int i = 0; i <= lastRowNumber; i++) {
            Cell cellData = sheet.getRow(i).getCell(cell);
            valueOfString.add(cellData.toString());
        }

        workbook.close();
        fis.close();
        return valueOfString;
    }

    public void addHyperlinktoExcel(String filePathForJPG) throws IOException {
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(filePathforExcel);
        } catch (FileNotFoundException ignored) {
            fis = new FileInputStream(filePathforExcel);
        }
        Workbook workbook = WorkbookFactory.create(fis);
        Sheet sheet = workbook.getSheetAt(0);

        // Add the hyperlink
        CreationHelper helper = workbook.getCreationHelper();
        CellStyle hyperlinkStyle = workbook.createCellStyle();
        Font hyperlinkFont = workbook.createFont();
        hyperlinkFont.setUnderline(Font.U_SINGLE);
        hyperlinkFont.setColor(IndexedColors.BLUE.getIndex());
        hyperlinkStyle.setFont(hyperlinkFont);

        // Get the row number to write 'Download Attachment' hyperlink for the appropriate cell
        rowExcel++;
        System.out.println("rowExcel = " + rowExcel + " : added successfully");
        Row row = sheet.getRow(rowExcel);
        Cell cell = row.createCell(4); // add hyperlink to the 4th column
        cell.setCellValue("Download Attachment");

        // Compute relative path
        File excelFile = new File(filePathforExcel);
        File jpgFile = new File(filePathForJPG);
        String relativePath = excelFile.getParentFile().toPath().relativize(jpgFile.toPath()).toString();
        relativePath = relativePath.replace("\\", "/");

        // Set the hyperlink
        Hyperlink hyperlink = helper.createHyperlink(HyperlinkType.URL);
        hyperlink.setAddress(relativePath);

        // Apply hyperlink to the cell
        cell.setHyperlink(hyperlink);
        cell.setCellStyle(hyperlinkStyle);

        // Save the Excel file
        try (FileOutputStream outputStream = new FileOutputStream(filePathforExcel)) {
            workbook.write(outputStream);
            System.out.println("Hyperlink written to Excel successfully.");
        }
        workbook.close();
    }
}
