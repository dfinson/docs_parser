package parser;

import com.google.re2j.*;
import lombok.Data;
import lombok.SneakyThrows;
import lombok.val;
import lombok.var;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileInputStream;
import java.util.*;
import java.util.stream.Collectors;

@Data
public class WordDocumentParser {
    private File document;
    private String title;
    private Map<String, Integer> wordCount;

    public WordDocumentParser(File file){
        document = file;
        wordCount = new LinkedHashMap<>();
        processDocument();
    }

    private void processDocument() {
        var lines = extractLines();
        lines = setTitle(lines);
        doSpeakersWordCount(lines);
    }

    private void doSpeakersWordCount(List<String> lines) {
        val discussionLines = extractActualDiscourse(lines);
        for (int i = 0; i < discussionLines.size(); i++) {
            val line = discussionLines.get(i);
            if(isStartOfSpeechSegment(line)){
                val speaker = line.replaceAll(":", "");
                wordCount.putIfAbsent(speaker, 0);
                val currentSpeechSegmentBuilder = new StringBuilder();
                while(isContinuationOfSpeechSegment(discussionLines, i)){
                    currentSpeechSegmentBuilder.append(discussionLines.get(i + 1));
                    i++;
                }
                int segmentWordCount = currentSpeechSegmentBuilder.toString().split("\\s+").length;
                wordCount.put(speaker, wordCount.get(speaker) + segmentWordCount);
            }
        }
    }

    private boolean isContinuationOfSpeechSegment(List<String> discussionLines, int i) {
        if(i + 1 >= discussionLines.size())
            return false;
        val line = discussionLines.get(i + 1);
        return !line.endsWith(":") || !isStartOfSpeechSegment(line);
    }

    private final static Pattern SEGMENT_START =
            com.google.re2j.Pattern.compile("(\\s*[\u05D0-\u05EA |\"| \\-]+\\s*)+\\s*:\\s*");
    private boolean isStartOfSpeechSegment(String line) {
        return SEGMENT_START.matcher(line).matches();
    }

    private List<String> extractActualDiscourse(List<String> lines) {
        val fromStart = nextStartOfYoshevRoshSpeaking(nextStartOfYoshevRoshSpeaking(lines));
        for (int i = 0; i < fromStart.size(); i++)
            if (fromStart.get(i).contains("הישיבה ננעלה"))
                return fromStart.subList(0, i);
        return fromStart;
    }

    private List<String> setTitle(List<String> lines) {
        String protocolNumber = null, protocolName = null;
        int i = 0;
        for (; i < Math.min(7, lines.size()) && protocolNumber == null; i++)
            if (lines.get(i).contains("פרוטוקול מס'"))
                protocolNumber = lines.get(i);
        for (; i < Math.min(15, lines.size()) && protocolName == null; i++) {
            val withoutSpacesOrHyphens = lines.get(i).replaceAll("\\s+", "").replaceAll("-", "");
            if (withoutSpacesOrHyphens.contains("סדרהיום")) {
                val titleNameBuilder = new StringBuilder(
                        lines.get(i).replaceAll("סדר היום:", "").replaceAll("סדר-היום:", "")
                );
                i++;
                while (i < lines.size() && !lines.get(i).contains("נכחו:")) {
                    titleNameBuilder.append(lines.get(i));
                    i++;
                }
                protocolName = titleNameBuilder.toString();
            }
        }
        title = "\t\tfile name: " + document.getName() + NEW_LINE +
                "\t\tprotocol number: " + nullSafeStringValue(protocolNumber) + NEW_LINE +
                "\t\tprotocol title: " + nullSafeStringValue(protocolName) + NEW_LINE +
                "\t\tword count by speaker: " + NEW_LINE + NEW_LINE;
        return i <= lines.size() ? lines.subList(i, lines.size()) : lines;
    }

    private String nullSafeStringValue(String str){
        if(str == null || str.equals("null") || str.length() > 500)
            return "Cannot be determined due to inconsistent spelling, format, or formatting";
        return str;
    }

    @SneakyThrows
    private List<String> extractLines() {
        val fileInputStream = new FileInputStream(document);
        val doc = new XWPFDocument(fileInputStream);
        val wordExtractor = new XWPFWordExtractor(doc);
        val rawText = wordExtractor.getText();
        return Arrays
                .stream(rawText.replaceAll("\t", "").split("\n"))
                .filter(line -> !line.equals(""))
                .map(String::trim)
                .collect(Collectors.toList());
    }

    private List<String> nextStartOfYoshevRoshSpeaking(List<String> lines){
        for (int i = 0; i < lines.size(); i++)
            if (lines.get(i).contains("יו\"ר") && lines.get(i).endsWith(":"))
                return lines.subList(i, lines.size());
        return lines;
    }

    private final static String NEW_LINE = System.lineSeparator();

}
