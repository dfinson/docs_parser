package parser;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Map;

@Data
@NoArgsConstructor
public class ParsedDocumentData {
    private String filename;
    private String title;
    private String number;
    private Map<String, Integer> wordCountBySpeaker;
}
