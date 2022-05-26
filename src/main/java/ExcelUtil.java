import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.util.Date;
import java.util.List;

public class ExcelUtil {
    /**
     * Method that creates an excel file, with headerTitles passed as an argument and data from a list of objects.
     * headerTitle at position 0 retrieves data from field at header 0...
     * @param headers field mapping for header - object
     * @param headersTitle headerTitles for excel file columns
     * @param objects a list of objects that contains the data we need to parse to excel
     * @throws Exception when length of header and headerTitles does not match
     */
    public static void createExcel(String headers, String headersTitle, List<?> objects) throws Exception {
        /**split headers and header titles by , to retrtieve titles and fields to map values */
        String[] headerColumns = headers.replaceAll("\\s+", "").split(",");
        String[] headerTitleColumns = headersTitle.replaceAll("\\s+", "").split(",");
        if (headerColumns.length != headerTitleColumns.length) {
            throw new Exception("Headers and Header Titles length is not equal");
        }
        /** create excel */
        Workbook workbook = new XSSFWorkbook();
        CreationHelper createHelper = workbook.getCreationHelper();
        Sheet sheet = workbook.createSheet("Export");
        sheet.setColumnWidth(0, 6000);
        sheet.setColumnWidth(1, 4000);

        /** Headers */
        Row header = sheet.createRow(0);

        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        CellStyle dateStyle = workbook.createCellStyle();
        dateStyle.setDataFormat(
                createHelper.createDataFormat().getFormat("dd-MM-yyyy"));

        for (int i = 0; i < headerTitleColumns.length; i++) {
            Cell headerCell = header.createCell(i);
            headerCell.setCellValue(headerTitleColumns[i]);
            headerCell.setCellStyle(headerStyle);
        }

        /** Data */
        for (int i = 0; i < objects.size(); i++) {
            Row row = sheet.createRow(i + 1);
            ObjectMapper objectMapper = new ObjectMapper();
            objectMapper.configure(SerializationFeature.FAIL_ON_EMPTY_BEANS, false);
            String objectToString = objectMapper.writeValueAsString(objects.get(i));
            for (int j = 0; j < headerColumns.length; j++) {
                Cell cell = row.createCell(j);
                /** split field by . for inner objects ex. onlineApplication.commonOnlineApplication.description */
                String[] path = headerColumns[j].split("\\.");
                int start = 0;
                for (int k = 0; k < path.length; k++) {
                    start = objectToString.indexOf(path[k], start);
                }
                /** Calculate start of value by finding index of field + field length + 2(":) */
                start = start + path[path.length - 1].length() + "\":".length();
                /** Calculate end of value by finding index of , after start index.Also check and remove object break character (}) */
                int end = objectToString.indexOf(",", start);
                if (objectToString.charAt(end - 1) == '}') {
                    end--;
                }
                String value = objectToString.substring(start, end);
                if (value.equals("null")) {
                    cell.setBlank();
                } else {
                    if (path[path.length - 1].contains("Date") || path[path.length - 1].contains("date")) {
                        Date date = Date.from(Instant.ofEpochMilli(Long.parseLong(value)));
                        cell.setCellStyle(dateStyle);
                        cell.setCellValue(date);
                    } else {
                        cell.setCellValue(value);
                    }
                }
            }
        }
        /** construct file */
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        workbook.write(baos);
        Files.write(Path.of("./myExcelExport"), baos.toByteArray());
        workbook.close();
        baos.close();
    }
}
