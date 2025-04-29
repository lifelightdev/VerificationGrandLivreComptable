import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import java.util.HashMap;
import java.util.Map;

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
        Assertions.assertEquals(account, result.account());
        Assertions.assertEquals(label, result.label());
    }

    @ParameterizedTest
    @CsvSource({
            "40100-0001 ORANGE, true",
            "C'est le nom du Syndic, false",
            "8 AVENUE DES CHAMPS ELYSE 14/04/2025 Page : 10, false",
            "75000 PARIS | , false",
            "001 NOM COPROPRIÉTÉ au 31/12/2024 Gestionnaire : NOM PRENOM, false"
    })
    public void extractIsLineAccount(String line, boolean is) {
        boolean result = ExtractInfo.isAcccount(line, "75000", "001");
        if (is) {
            Assertions.assertTrue(result);
        } else {
            Assertions.assertFalse(result);
        }
    }

    @ParameterizedTest
    @CsvSource({
            "01/01/2024 40100-0001 Report de 0.00 € 3 210.69 € 3 210.69 €,              '',    01/01/2024, 40100-0001, '', '',    '',   Report de 0.00€,       3 210.69€,  3 210.69€",
            "01/01/2024 40100-0002 Report de 0.00 € 432.93 € 432.93 €,                  '',    01/01/2024, 40100-0002, '', '',    '',   Report de 0.00€,         432.93€,    432.93€",
            "01/01/2024 40100-0003 Report de -1 234.56 € 23 456.78 € 24 691.34 €,       '',    01/01/2024, 40100-0003, '', '',    '',   Report de -1 234.56€, 23 456.78€, 24 691.34€",
            "33333 01/01/2024 10500 15 44444 APPEL FONDS LOI ALUR  2 000.00 €,          33333, 01/01/2024, 10500,      15, 44444, '',   APPEL FONDS LOI ALUR,         '',  2 000.00€",
            "111111 01/01/2024 40100-0001 | VI 55555 Virt HONORAIRE COURANT 3 000.00 €, 11111, 01/01/2024, 40100-0001, VI, 55555, Virt, HONORAIRE COURANT,     3 000.00€,          ''",
            "01/01/2024 40100-0001 Report de 0.00 € 100 000.00€| 100 000.00 €,             '', 01/01/2024, 40100-0001, '',    '', '',   Report de 0.00€,     100 000.00€, 100 000.00€"
    })
    public void extractline(String line, String document, String date, String account, String journal,
                            String counterpart, String checkNumber, String label, String debit, String credit) {
        Map<String, Account> accounts = new HashMap<>();
        accounts.put("40100-0001", new Account("40100-0001", "Orange"));
        accounts.put("40100-0002", new Account("40100-0002", "EDF"));
        accounts.put("40100-0003", new Account("40100-0003", "TOTAL"));
        accounts.put("10500", new Account("10500", "Fond travaux"));
        Line result = ExtractInfo.line(line, accounts);
        Assertions.assertNotNull(result);
        Assertions.assertEquals(document, result.document());
        Assertions.assertEquals(date, result.date());
        Assertions.assertEquals(account, result.account().account());
        Assertions.assertEquals(journal, result.journal());
        Assertions.assertEquals(counterpart, result.counterpart());
        Assertions.assertEquals(checkNumber, result.checkNumber());
        Assertions.assertEquals(label, result.label());
        Assertions.assertEquals(debit, result.debit());
        Assertions.assertEquals(credit, result.credit());
    }
}
