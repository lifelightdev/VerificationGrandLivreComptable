import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

public class ExtractInfoTest {
    @Test
    public void extractSyndicName() {
        String ligne = "C'est le nom du Syndic";
        String name = ExtractInfo.syndicName(ligne);
        Assertions.assertEquals(ligne, name);
    }
}
