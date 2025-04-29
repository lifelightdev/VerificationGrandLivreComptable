import life.light.Account;
import life.light.ExtractInfo;
import life.light.Line;
import life.light.TotalAccount;
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
            "10240 TRAVAUX PORTE PARKING,                                         10240, TRAVAUX PORTE PARKING",
            "40100-0001 ORANGE,                                              40100-0001, ORANGE",
            "40100-0002 _ RELANCE,                                           40100-0002, RELANCE",
            "40100-0609 | VBP HUISSIERS DE JUSTICE,                          40100-0609, VBP HUISSIERS DE JUSTICE",
            "45000-0003 — TRUC MUCHE,                                        45000-0003, TRUC MUCHE",
            "40100-0027 … RENOV,                                             40100-0027, RENOV",
            "40100-0890 | FBI - FUITE BATIMENT INVESTIGA°,                   40100-0890, FBI - FUITE BATIMENT INVESTIGA",
            "45000-0002 CHRISTOPHE******,                                    45000-0002, CHRISTOPHE",
            "45000-0000 DUPONT / DUPOND # ****** 4% 4444445 5KKEEAXKRRRRRHE, 45000-0000, DUPONT / DUPOND"
})
    public void extractAccount(String line, String account, String label) {
        Account result = ExtractInfo.account(line);
        Assertions.assertEquals(account, result.account());
        Assertions.assertEquals(label, result.label());
    }

    @ParameterizedTest
    @CsvSource({
            "40100-0001 ORANGE,                                                 true",
            "C'est le nom du Syndic,                                            false",
            "8 AVENUE DES CHAMPS ELYSE 14/04/2025 Page : 10,                    false",
            "75000 PARIS | ,                                                    false",
            "001 NOM COPROPRIÉTÉ au 31/12/2024 Gestionnaire : NOM PRENOM,       false"
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
            "01/01/2024 40100-0001 Report de 0.00 € 100 000.00€| 100 000.00 €,             '', 01/01/2024, 40100-0001, '',    '', '',   Report de 0.00€,     100 000.00€, 100 000.00€",
            "|44444 01/10/2024 40100-0001 40 10500 TRAVAUX PORTE PARKING  10 749.26 €|, 44444, 01/10/2024, 40100-0001, 40, 10500, '',   TRAVAUX PORTE PARKING,        '',  10 749.26€",
            "01/01/2024 40100-0001 Report de 0.00 € / 440.12 € / 440.12 €,                 '', 01/01/2024, 40100-0001, '',    '', '',   Report de 0.00€,         440.12€,     440.12€"
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

    @ParameterizedTest
    @CsvSource({
            "40100-0001 ORANGE,                                                 false",
            "C'est le nom du Syndic,                                            false",
            "8 AVENUE DES CHAMPS ELYSE 14/04/2025 Page : 10,                    false",
            "75000 PARIS | ,                                                    false",
            "001 NOM COPROPRIÉTÉ au 31/12/2024 Gestionnaire : NOM PRENOM,       false",
            "01/01/2024 40100-0001 Report de 0.00 € 3 210.69 € 3 210.69 €,      true",
            "33333 01/01/2024 10500 15 44444 APPEL FONDS LOI ALUR  2 000.00 € , true"
    })
    public void extractIsLine(String line, boolean is) {
        boolean result = ExtractInfo.isLigne(line);
        if (is) {
            Assertions.assertTrue(result);
        } else {
            Assertions.assertFalse(result);
        }
    }

    @ParameterizedTest
    @CsvSource({
            "Total compte 40100-0001 (Solde : 0.00 €) 100 000.00 € 100 000.00 €, Total compte 40100-0001 (Solde : 0.00€),                40100-0001, 100 000.00€, 100 000.00€",
            "Total compte 40100-0002 (Solde : 0.00 €) 999 999.99 € 999 999.99 €, Total compte 40100-0002 (Solde : 0.00€),                40100-0002, 999 999.99€, 999 999.99€",
            "Total compte 40100-0002 (Solde : 0.00 €) 1 238.40 € ,               Total compte 40100-0002 (Solde : 0.00€),                40100-0002,   1 238.40€,   1 238.40€",
            "Total compte 40100-0001 (Solde créditeur : -3 896.22 €),            Total compte 40100-0001 (Solde créditeur : -3 896.22€), 40100-0001,          '',           '' "
    })
    public void extractTotalAccount(String line, String label, String account, String debit, String credit) {
        Map<String, Account> accounts = new HashMap<>();
        accounts.put("40100-0001", new Account("40100-0001", "Orange"));
        accounts.put("40100-0002", new Account("40100-0002", "EDF"));
        TotalAccount result = ExtractInfo.totalAccount(line, accounts);
        Assertions.assertNotNull(result);
        Assertions.assertEquals(label, result.label());
        Assertions.assertEquals(account, result.account().account());
        Assertions.assertEquals(debit, result.debit());
        Assertions.assertEquals(credit, result.credit());
    }

    @ParameterizedTest
    @CsvSource({
            "40100-0001 ORANGE,                                                  false",
            "C'est le nom du Syndic,                                             false",
            "8 AVENUE DES CHAMPS ELYSE 14/04/2025 Page : 10,                     false",
            "75000 PARIS | ,                                                     false",
            "001 NOM COPROPRIÉTÉ au 31/12/2024 Gestionnaire : NOM PRENOM,        false",
            "01/01/2024 40100-0001 Report de 0.00 € 3 210.69 € 3 210.69 €,       false",
            "33333 01/01/2024 10500 15 44444 APPEL FONDS LOI ALUR  2 000.00 € ,  false",
            "Total compte 40100-0001 (Solde : 0.00 €) 100 000.00 € 100 000.00 €, true"
    })
    public void extractIsTotalAccount(String line, boolean is) {
        boolean result = ExtractInfo.isTotalAccount(line);
        if (is) {
            Assertions.assertTrue(result);
        } else {
            Assertions.assertFalse(result);
        }
    }

}
