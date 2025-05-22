package life.light.extract.info;

import life.light.FileOfTest;
import life.light.type.*;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import java.util.HashMap;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;

class LedgerTest {

    Map<String, TypeAccount> accounts = new HashMap<>();
    TypeAccount accountBank1 = new TypeAccount("51220", "Banque 1");
    TypeAccount accountBank2 = new TypeAccount("51221", "Banque 2");
    String nameFileTestLedger = "";

    @BeforeEach
    void setUp() {
        accounts.put(accountBank1.account(), accountBank1);
        accounts.put(accountBank2.account(), accountBank2);
        FileOfTest fileOfTest = new FileOfTest();
        try {
            nameFileTestLedger = fileOfTest.createMinimalLedgerFile();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @ParameterizedTest
    @CsvSource({
            "C'est le nom du Syndic, C'est le nom du Syndic",
            "C'est le nom du Syndic|, C'est le nom du Syndic"
    })
    void extractSyndicName(String line, String SyndicName) {
        String name = Ledger.syndicName(line);
        Assertions.assertEquals(SyndicName, name);
    }

    @ParameterizedTest
    @CsvSource({
            "8 AVENUE DES CHAMPS ELYSE 11/04/2025 Page : 1,  11/04/2025",
            "8 AVENUE DES CHAMPS ELYSE 14/04/2025 Page : 10, 14/04/2025"
    })
    void extractPrintDate(String line, String printDate) {
        String date = Ledger.printDate(line);
        Assertions.assertEquals(printDate, date);
    }

    @ParameterizedTest
    @CsvSource({
            "01.01.01.01.01 Grand Livre arrêté au 31/12/2024,  31/12/2024",
            "02.02.02.02.02 Grand Livre arrêté au 31/12/2023,  31/12/2023"
    })
    void extractStopDate(String line, String stopDate) {
        String date = Ledger.stopDate(line);
        Assertions.assertEquals(stopDate, date);
    }

    @ParameterizedTest
    @CsvSource({
            "75000 PARIS Grand Livre arrêté au 31/12/2024, 75000"
    })
    void extractPostalCode(String line, String postalCode) {
        String code = Ledger.postalCode(line);
        Assertions.assertEquals(postalCode, code);
    }

    @Test
    void getInfoGrandLivre() {
        InfoGrandLivre infoGrandLivre = Ledger.getInfoGrandLivre(nameFileTestLedger);
        assertEquals("C'est le nom du Syndic", infoGrandLivre.syndicName());
        assertEquals("11/04/2025", infoGrandLivre.printDate());
        assertEquals("2024-12-31", infoGrandLivre.stopDate().toString());
        assertEquals("75000", infoGrandLivre.postalCode());
    }

    @Test
    void getNumberOfLineInFile() {
        int result = Ledger.getNumberOfLineInFile(nameFileTestLedger);
        assertEquals(11, result);
    }

    @ParameterizedTest
    @CsvSource({
            "10240 TRAVAUX PARKING,                                         10240, TRAVAUX PARKING",
            "40100-0001 ORANGE,                                              40100-0001, ORANGE",
            "40100-0002 _ RELANCE,                                           40100-0002, RELANCE",
            "40100-0609 | VBP HUISSIERS DE JUSTICE,                          40100-0609, VBP HUISSIERS DE JUSTICE",
            "45000-0003 — TRUC MUCHE,                                        45000-0003, TRUC MUCHE",
            "40100-0001 — NOM du compte,                                     40100-0001, NOM du compte",
            "40100-0027 … RENOV,                                             40100-0027, RENOV",
            "40100-0890 | INVESTIGAT°,                                       40100-0890, INVESTIGAT",
            "45000-0002 CHRISTOPHE******,                                    45000-0002, CHRISTOPHE",
            "45000-0000 DUPONT / DUPOND # ****** 4% 4444445 5KKEEAXKRRRRRHE, 45000-0000, DUPONT / DUPOND",
            "461VC VENDEURS CREDITEURS,                                           461VC, VENDEURS CREDITEURS",
            "45000-0001 DUPONT 31/12/2020,                                   45000-0001, DUPONT 31/12/2020",
            "46200 VIR TRUC 30/03/2021 NON IDENTIFIE,                             46200, VIR TRUC 30/03/2021 NON IDENTIFIE",
            "4A0100-0077 | CHRISTAL,                                         40100-0077, CHRISTAL"
    })
    void account(String line, String account, String label) {
        TypeAccount result = Ledger.account(line);
        Assertions.assertEquals(account, result.account());
        Assertions.assertEquals(label, result.label());
    }

    @ParameterizedTest(name = "Given {0} return {1}")
    @CsvSource({
            "40100-0001 ORANGE,                                                 true",
            "C'est le nom du Syndic,                                            false",
            "8 AVENUE DES CHAMPS ELYSE 14/04/2025 Page : 10,                    false",
            "8 AVENUE DES CHAMPS ELYSE 11/04/2025 Page: 63,                     false",
            "75000 PARIS | ,                                                    false",
            "001 NOM COPROPRIÉTÉ au 31/12/2024 Gestionnaire : NOM PRENOM,       false",
            "45000-0001 DUPONT 31/12/2020,                                      true",
            "461VC VENDEURS CREDITEURS,                                         true",
            "4A0100-0077 | CHRISTAL,                                            true",
            "43575  01/01/2024 51220 OD 45000 VOTRE VIREMENT 100.00€,           false",
            "42206  26/06/2024 70246  P4 45020  APRT     JARD,                  false",
            "1 238.40 €,                                                        false"
    })
    void isAccount(String line, boolean is) {
        boolean result = Ledger.isAcccount(line, "75000", "001");
        if (is) {
            Assertions.assertTrue(result);
        } else {
            Assertions.assertFalse(result);
        }
    }

    @ParameterizedTest
    @CsvSource({
            "01/01/2024 40100-0001 Report de 0.00 € 3 210.69 € 3 210.69 €,                      '', 01/01/2024, 40100-0001, '',    '',     '',   Report de 0.00€,         3210.69,   3210.69",
            "01/01/2024 40800  Report de 0.00 €  107 551.98 €  107 551.98 €,                    '', 01/01/2024, 40800,      '',    '',     '',   Report de 0.00€,       107551.98, 107551.98",
            "01/01/2024 40100-0002 Report de 0.00 € 432.93 € 432.93 €,                          '', 01/01/2024, 40100-0002, '',    '',     '',   Report de 0.00€,          432.93,    432.93",
            "01/01/2024 40100-0003 Report de -1 234.56 € 23 456.78 € 24 691.34 €,               '', 01/01/2024, 40100-0003, '',    '',     '',   Report de -1 234.56€,   23456.78,  24691.34",
            "33333  01/01/2024 10500 15 51220 APPEL FONDS LOI ALUR  2 000.00 €,              33333, 01/01/2024, 10500,      15, 51220,     '',   APPEL FONDS LOI ALUR,         '',   2000.00",
            "'111111 01/01/2024 40100-0001 | VI 10500 Virt HONORAIRE COURANT 3 000.00 € ',   11111, 01/01/2024, 40100-0001, VI, 10500,   Virt, HONORAIRE COURANT,       3000.00,        ''",
            "01/01/2024 40100-0001 Report de 0.00 € 100 000.00€| 100 000.00 €,                  '', 01/01/2024, 40100-0001, '',    '',     '',   Report de 0.00€,       100000.00, 100000.00",
            "|44444 01/10/2024 40100-0001 40 10500 TRAVAUX PARKING  10 749.26 €|,            44444, 01/10/2024, 40100-0001, 40, 10500,     '',   TRAVAUX PARKING,              '',  10749.26",
            "01/01/2024 40100-0001 Report de 0.00 € / 440.12 € / 440.12 €,                      '', 01/01/2024, 40100-0001, '',    '',     '',   Report de 0.00€,          440.12,    440.12",
            "01/01/2024 40100-0003 Report de -2 592.52 €  2592.52 €,                            '', 01/01/2024, 40100-0003, '',    '',     '',   Report de -2 592.52€,         '',   2592.52",
            "'33333 01/01/2024 45010-0001 20 10500 APPEL DE FONDS 44961 € ',                 33333, 01/01/2024, 45000-0001, 20, 10500,     '',   APPEL DE FONDS,           449.61,        ''",
            "|88888 16/04/2024 40100-0002 VI 51220 Virt EDF 08/04/24-28/04/24 1800.00 € l,   88888, 16/04/2024, 40100-0002, VI, 51220,   Virt, EDF 08/04/24-28/04/24,   1800.00,        ''",
            "'44917 29/03/2024 40100-005 7 AC 51220 PARIS 03/2024 64.89 € ',                 44917, 29/03/2024, 40100-0057, AC, 51220,     '',   PARIS 03/2024,             64.89,        ''",
            "'35138 29/01/2024 40100-0001 VI 51220 Virt Orange 2024 2 361.62 € ',            35138, 29/01/2024, 40100-0001, VI, 51220,    Virt, Orange 2024,             2361.62,        ''",
            "144796 05/09/2024 40100-0001 — AC 51220 Orange  1 599.40 €|,                    44796, 05/09/2024, 40100-0001, AC, 51220,      '',   Orange,                       '',   1599.40",
            "144796 27/09/2024 40100-0001 . VI 51220 Virt Orange 1 599.40 € |,               44796, 27/09/2024, 40100-0001, VI, 51220,    Virt, Orange,                  1599.40,        ''",
            "'140040 30/04/2024 40100-0057OD 51220  PRLVT ORANGE 84.89 € ',                  40040, 30/04/2024, 40100-0057, OD, 51220,      '', PRLVT ORANGE,              84.89,        ''",
            "'45476 01/10/2024 40800  40 45020  TRAVAUX PARKING  10 749.26 € ',              45476, 01/10/2024, 40800,      40, 45020,      '', TRAVAUX PARKING,        10749.26,        ''",
            "'55617 31/12/2024 40800  OD 45020 TRAVAUX PARKING  10 749.26 €',                55617, 31/12/2024, 40800,      OD, 45020,      '', TRAVAUX PARKING,              '',  10749.26",
            "41751 01/07/2024 40100-0001  OD 51220  HONORAIRES 3T2024 (00001806) 3 000.00 €, 41751, 01/07/2024, 40100-0001, OD, 51220,      '', HONORAIRES 3T2024 (00001806), '',   3000.00",
            "'42201  01/06/2024 45000-0001 P2 10500  APUREMENT CHARGES 2 312.66 € ',         42201, 01/06/2024, 45000-0001, P2, 10500,      '', APUREMENT CHARGES,       2312.66,        ''",
            "55555 01/01/2024 10500 P4 40800 TRAVAUX,                                        55555, 01/01/2024, 10500,      P4, 40800,      '', TRAVAUX,                      '',        ''",
            "'34257 05/01/2024 45010-0001 OD 40100-0001 HONOPRAIRES 120.00 € ',              34257, 05/01/2024, 45000-0001, OD, 40100-0001, '', HONOPRAIRES,              120.00,        ''",
            "36144 01/01/2024 51220 OD 431000001 URSSAF 4 000.00 €,                          36144, 01/01/2024, 51220,      OD, 43100,      '', URSSAF,                       '',   4000.00"
    })
    void line(String line, String document, String date, String account, String journal,
              String counterpart, String checkNumber, String label, String debit, String credit) {
        Map<String, TypeAccount> accounts = new HashMap<>();
        accounts.put("10500", new TypeAccount("10500", "Fond travaux"));
        accounts.put("40100-0001", new TypeAccount("40100-0001", "Orange"));
        accounts.put("40100-0002", new TypeAccount("40100-0002", "EDF"));
        accounts.put("40100-0003", new TypeAccount("40100-0003", "TOTAL"));
        accounts.put("40100-0057", new TypeAccount("40100-0057", "PARIS"));
        accounts.put("43100", new TypeAccount("43100", "Compte"));
        accounts.put("40800", new TypeAccount("40800", "Un compte"));
        accounts.put("45000-0001", new TypeAccount("45000-0001", "Monsieur DUPONT"));
        accounts.put("51220", new TypeAccount("51220", "Banque"));
        Line result = Ledger.line(line, accounts);
        Assertions.assertNotNull(result);
        Assertions.assertEquals(document, result.document());
        Assertions.assertEquals(date, result.date());
        Assertions.assertEquals(account, result.account().account());
        Assertions.assertEquals(journal, result.journal());
        if ("".equals(counterpart)) {
            Assertions.assertNull(result.accountCounterpart());
        } else {
            Assertions.assertEquals(counterpart, result.accountCounterpart().account());
        }
        Assertions.assertEquals(checkNumber, result.checkNumber());
        Assertions.assertEquals(label, result.label());
        Assertions.assertEquals(debit, result.debit());
        Assertions.assertEquals(credit, result.credit());
    }

    @ParameterizedTest
    @CsvSource({
            "40100-0001 ORANGE,                                                               false",
            "C'est le nom du Syndic,                                                          false",
            "8 AVENUE DES CHAMPS ELYSE 14/04/2025 Page : 10,                                  false",
            "75000 PARIS | ,                                                                  false",
            "001 NOM COPROPRIÉTÉ au 31/12/2024 Gestionnaire : NOM PRENOM,                     false",
            "01/01/2024  40100-0001 Report de 0.00 € 3 210.69 € 3 210.69 €,                   true",
            "33333 01/01/2024 10500 15 44444 APPEL FONDS LOI ALUR  2 000.00 € ,               true",
            "01.01.01.01.01,                                                                  false",
            "41751 01/07/2024 40100-0001  OD 62100  HONORAIRES 3T2024 (00001806)  3 000.00 €, true",
            "'42201  01/06/2024 45000-0001 P2 10500  APUREMENT CHARGES 2 312.66 € ',          true"
    })
    void isLine(String line, boolean is) {
        boolean result = Ledger.isLigne(line);
        if (is) {
            Assertions.assertTrue(result);
        } else {
            Assertions.assertFalse(result);
        }
    }


    @ParameterizedTest
    @CsvSource({
            "Total compte 40100-0001 (Solde : 0.00 €) 100 000.00 € 100 000.00 €, Total compte (Solde : 0.00€),                40100-0001, 100000.00, 100000.00",
            "Total compte 40100-0002 (Solde : 0.00 €) 999 999.99 € 999 999.99 €, Total compte (Solde : 0.00€),                40100-0002, 999999.99, 999999.99",
            "Total compte 40100-0002 (Solde : 0.00 €) 1 238.40 € ,               Total compte (Solde : 0.00€),                40100-0002,   1238.40,   1238.40",
            "Total compte 40100-0001 (Solde créditeur : -3 896.22 €),            Total compte (Solde créditeur : -3 896.22€), 40100-0001,        '',        ''",
            "Total compte 40100-0001 (Solde : 0.00 €) 16144.77€| — 16144.77€,    Total compte (Solde : 0.00€),                40100-0001,  16144.77,  16144.77",
            "Total compte 40100-001 7 (Solde : 0.00 €) 2 377.01 € 2 377.01 €,    Total compte (Solde : 0.00€),                40100-0017,   2377.01,   2377.01"
    })
    void extractTotalAccount(String line, String label, String account, String debit, String credit) {
        Map<String, TypeAccount> accounts = new HashMap<>();
        accounts.put("40100-0001", new TypeAccount("40100-0001", "Orange"));
        accounts.put("40100-0002", new TypeAccount("40100-0002", "EDF"));
        accounts.put("40100-0017", new TypeAccount("40100-0017", "GDF"));
        TotalAccount result = Ledger.totalAccount(line, accounts);
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
    void extractIsTotalAccount(String line, boolean is) {
        boolean result = Ledger.isTotalAccount(line);
        if (is) {
            Assertions.assertTrue(result);
        } else {
            Assertions.assertFalse(result);
        }
    }

    @ParameterizedTest
    @CsvSource({
            "Total immeuble (Solde : 0.00 €) 100 000.00 € 100 000.00 €,true",
            " (Solde : 0.00 €) 10.00 € 100 000.00 €                   ,false",
    })
    void isTotalBuilding(String line, boolean is) {
        boolean result = Ledger.isTotalBuilding(line);
        if (is) {
            Assertions.assertTrue(result);
        } else {
            Assertions.assertFalse(result);
        }
    }

    @ParameterizedTest
    @CsvSource({
            "Total immeuble (Solde : 0.00 €) 100 000.00 € 100 000.00 €, Total immeuble (Solde : 0.00€), 100000.00, 100000.00",
    })
    void totalBuilding(String line, String label, String debit, String credit) {
        TotalBuilding result = Ledger.totalBuilding(line);
        Assertions.assertNotNull(result);
        Assertions.assertEquals(label, result.label());
        Assertions.assertEquals(debit, result.debit());
        Assertions.assertEquals(credit, result.credit());
    }
}