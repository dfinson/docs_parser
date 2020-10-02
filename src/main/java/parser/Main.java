package parser;

import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import lombok.val;
import lombok.var;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.*;
import java.nio.file.Paths;
import java.util.*;
import java.util.stream.Collectors;

public class Main {
    public static void main(String[] args){
        val rootDir = Paths.get(args != null && args.length > 0 ? args[0] : "").toAbsolutePath();
        System.out.println("starting parse of root directory " + rootDir.toString());
        val dirParser = new DirectoryParser(rootDir);
        System.out.println("completed parse of root directory " + rootDir.toString());
        System.out.println("totaled " + dirParser.getSpeakersWordCount().size() + " files");
        //printToTxtFile(dirParser.getSpeakersWordCount());
        printToExcelSpreadsheet(dirParser.getParsedDocumentDataList());
    }

    @SneakyThrows
    private static void printToExcelSpreadsheet(List<ParsedDocumentData> parsedDocumentDataList){
        val workbook = new XSSFWorkbook();
        for(val docData : parsedDocumentDataList){

            int row = 0;
            val filename = docData.getFilename();
            val protocolTab = sheet(workbook, filename);

            val firstRow = protocolTab.createRow(row++);
            val protocolNameCell = firstRow.createCell(0);
            protocolNameCell.setCellValue(docData.getTitle());
            protocolTab.addMergedRegion(new CellRangeAddress(row, row, 0,1));

            val secondRow = protocolTab.createRow(row++);
            val protocolNumberCell = secondRow.createCell(0);
            protocolNumberCell.setCellValue(docData.getNumber());
            protocolTab.addMergedRegion(new CellRangeAddress(row, row, 0,1));

            for (val entry : docData.getWordCountBySpeaker().entrySet()){
                val speaker = entry.getKey();
                val wordCount = entry.getValue();
                val currentRow = protocolTab.createRow(++row);
                val speakerCell = currentRow.createCell(1);
                speakerCell.setCellValue(speaker);
                val wordCountCell = currentRow.createCell(0);
                wordCountCell.setCellValue(wordCount);
            }
            protocolTab.autoSizeColumn(0, true);
            protocolTab.autoSizeColumn(1, true);
        }
        FileOutputStream outputStream = new FileOutputStream("result.xlsx");
        workbook.write(outputStream);
    }

    private static XSSFSheet sheet(XSSFWorkbook workbook, String filename){
        return sheet(workbook, filename, 0);
    }
    private static XSSFSheet sheet(XSSFWorkbook workbook, String filename, int depth){
        if(workbook.getSheet(filename) == null)
            return workbook.createSheet(filename);
        return sheet(workbook, filename + "(" + depth + ")", ++depth);
    }

    @SneakyThrows
    private static void printToTxtFile(Map<String, Map<String, Integer>> speakersWordCount) {
        File statText = new File("result.txt");
        FileOutputStream is = new FileOutputStream(statText);
        OutputStreamWriter osw = new OutputStreamWriter(is);
        Writer w = new BufferedWriter(osw);
        val result = new StringBuilder();
        int i = 0;
        for (Map.Entry<String, Map<String, Integer>> entry : speakersWordCount.entrySet()) {
            String title = entry.getKey();
            Map<String, Integer> speakersCount = entry.getValue();
            result
                    .append(System.lineSeparator())
                    .append("\t").append(++i).append(".").append(System.lineSeparator()).append(System.lineSeparator())
                    .append(title);
            speakersCount.forEach((speaker, wordCount) -> {
                result
                        .append("\t\t\t")
                        .append(speaker)
                        .append(": ")
                        .append(wordCount)
                        .append(System.lineSeparator());
            });
            if(speakersCount.isEmpty())
                result.append("\t\t\tCannot be determined due to either a corrupted document or a formatting inconsistency");
            result.append(System.lineSeparator());
        }
        w.write(result.toString());
        w.close();
    }
}
