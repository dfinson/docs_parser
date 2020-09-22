package parser;

import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import lombok.extern.slf4j.XSlf4j;
import lombok.val;

import java.io.File;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.logging.Logger;

@Data
public class DirectoryParser {

    private Map<String, Map<String, Integer>> speakersWordCount;

    public DirectoryParser(Path rootDir) {
        speakersWordCount = new LinkedHashMap<>();
        processDir(rootDir.toFile());
    }

    private void processDir(File entry) {
        if(entry.isDirectory() && entry.listFiles() != null) {
            System.out.println("parsing subdirectories of " + entry.toString());
            Arrays.stream(entry.listFiles()).forEach(this::processDir);
        } else if(isWordDocument(entry)) {
            System.out.println("parsing document " + entry.getName());
            val docParser = new WordDocumentParser(entry);
            speakersWordCount.put(docParser.getTitle(), docParser.getWordCount());
        }
    }

    private boolean isWordDocument(File entry) {
        return entry.isFile() && (entry.getName().endsWith(".docx") || entry.getName().endsWith(".doc"));
    }
}
