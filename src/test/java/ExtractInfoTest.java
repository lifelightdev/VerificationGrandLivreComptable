import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

public class ExtractInfoTest {

    @ParameterizedTest
    @CsvSource({
            "C'est le nom du Syndic, C'est le nom du Syndic",
            "C'est le nom du Syndic|, C'est le nom du Syndic"
    })
    public void extractSyndicName(String line, String SyndicName) {
        String name = ExtractInfo.syndicName(line);
        Assertions.assertEquals(SyndicName, name);
    }
}
