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

    @ParameterizedTest
    @CsvSource({
            "8 AVENUE DES CHAMPS ELYSE 11/04/2025 Page : 1,  11/04/2025",
            "8 AVENUE DES CHAMPS ELYSE 14/04/2025 Page : 10, 14/04/2025"
    })
    public void extractPrintDate(String line, String printDate) {
        String date = ExtractInfo.printDate(line);
        Assertions.assertEquals(printDate, date);
    }

    @ParameterizedTest
    @CsvSource({
            "01.01.01.01.01 Grand Livre arrêté au 31/12/2024,  31/12/2024",
            "02.02.02.02.02 Grand Livre arrêté au 31/12/2023,  31/12/2023"
    })
    public void extractStopDate(String line, String stopDate) {
        String date = ExtractInfo.stopDate(line);
        Assertions.assertEquals(stopDate, date);
    }

    @ParameterizedTest
    @CsvSource({
            "10240 TRAVAUX PORTE PARKING, 10240, TRAVAUX PORTE PARKING",
            "40100-0001 ORANGE, 40100-0001, ORANGE"
    })
    public void extractAccount(String line, String account, String label) {
        Account result = ExtractInfo.account(line);
        Assertions.assertEquals(account, result.account);
        Assertions.assertEquals(label, result.label);
    }

}
