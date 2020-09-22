package parser;

import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import lombok.val;
import lombok.var;
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
        exportResults(dirParser.getSpeakersWordCount());
    }

    @SneakyThrows
    private static void exportResults(Map<String, Map<String, Integer>> speakersWordCount) {
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
