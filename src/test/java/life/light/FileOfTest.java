package life.light;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.List;

public class FileOfTest {

    public final static Path tempTestDir = Paths.get("." + File.separator + "TEST_temp" + File.separator);

    public FileOfTest() {
        File temp = new File(tempTestDir.toString());
        temp.mkdir();
    }

    public String createMinimalLedgerFile() {
        try {
            List<String> lines = Arrays.asList(
                    "C'est le nom du Syndic        ",
                    "8 AVENUE DES CHAMPS ELYSE 11/04/2025 Page : 1",
                    "",
                    "75000 PARIS Grand Livre arrêté au 31/12/2024",
                    "",
                    "Pièce Date              Jal C-Partie  N° chèque Libellé                                      Débif      Crédit",
                    "",
                    "001 NOM COPROPRIÉTÉ au 31/12/2024 Gestionnaire : NOM PRENOM",
                    "",
                    "01.01.01.01.01",
                    "",
                    "40100-0001 — NOM du compte",
                    "      01/01/2024            40100-0001          Report de 0.00 €                             3 000.00 € 3 000.00 €",
                    "33333 01/01/2024            45010-0001          20 512 APPEL DE FONDS                        3 000.00 € ",
                    "                                                Total compte 40100-0001 (Solde : 3 000.00 €) 6 000.00 € 3 000.00 €",
                    " ",
                    " ",
                    "                                                Total immeuble (Solde : 0.00 €)              6 000.00 € 3 000.00 €",
                    ""
            );
            Files.write(tempTestDir.resolve("grand_livre_TEST.txt"), lines);
            return tempTestDir + File.separator + "grand_livre_TEST.txt";
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public String createMinimalLedgerFilePrintDateKO() {
        try {
            List<String> lines = Arrays.asList(
                    "C'est le nom du Syndic        ",
                    "8 AVENUE DES CHAMPS ELYSE 11/31/25 Page : 1",
                    "",
                    "75000 PARIS Grand Livre arrêté au 12/31/2O24",
                    "",
                    "Pièce Date              Jal C-Partie  N° chèque Libellé                                      Débif      Crédit",
                    "",
                    "001 NOM COPROPRIÉTÉ au 31/12/2024 Gestionnaire : NOM PRENOM",
                    "",
                    "01.01.01.01.01",
                    "",
                    "40100-0001 — NOM du compte",
                    "      01/01/2024            40100-0001          Report de 0.00 €                             3 000.00 € 3 000.00 €",
                    "33333 01/01/2024            45010-0001          20 512 APPEL DE FONDS                        3 000.00 € ",
                    "                                                Total compte 40100-0001 (Solde : 3 000.00 €) 6 000.00 € 3 000.00 €",
                    " ",
                    " ",
                    "                                                Total immeuble (Solde : 0.00 €)              6 000.00 € 3 000.00 €",
                    ""
            );
            Files.write(tempTestDir.resolve("grand_livre_TEST.txt"), lines);
            return tempTestDir + File.separator + "grand_livre_TEST.txt";
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public void copyBankFiles(String accountBank1, String accountBank2) {
        try {
            File tempBank = new File(tempTestDir + File.separator + "bank" + File.separator);
            tempBank.mkdir();
            File tempBank1 = new File(tempBank.getAbsoluteFile() + File.separator + accountBank1 + File.separator);
            tempBank1.mkdir();
            List<String> bank1Lines = Arrays.asList(
                    "BANQUE",
                    "RELEVE DE COMPTE COURANT",
                    "IBAN : ",
                    "BIC : ",
                    "31.10 ANCIEN SOLDE                                                                                      30500,90",
                    "",
                    "01.11 VIR INST M. TINTIN                                                 01.11.24                        1500,88",
                    "LIBELLE:Tintin ref : 1.018.0031",
                    "REF.CLIENT-NOTPROVIDED",
                    "",
                    "01.11 VIR SEPA DUPONT                                                    01.11.24                          400,00",
                    "LIBELLE:018 002 Dupont",
                    "REF.CLIENT:XXXXXXX",
                    "",
                    "04.11 VIR SEPA  MME DUPOND                                               04.11.24                          90,55",
                    "LIBELLE:APPEL DE FONDS 4T2024",
                    "REF.CLIENT:1 018 0033",
                    " ",
                    "15.11 VIR SEPA TOURNESOL                                                 15.11.24                        2000,00",
                    "REF.CLIENT:1-018-0075 Tournesol",
                    " ",
                    "30.11 VIR INST MONSIEUR HADOCK                                           30.11.24                           10,22",
                    " ",
                    "LIBELLE:Ref: 1-018-0032",
                    "REF.CLIENT:NOTPROVIDED",
                    " ",
                    "SOLDE INTERMEDIAIRE À FIN NOVEMBRE                                                                      34502,55",
                    " ",
                    " ",
                    "02.12  SOLDE EN EUROS                                                                                   34502,55"
            );
            Files.write(tempTestDir.resolve(tempBank1.getAbsoluteFile() + File.separator + accountBank1 + " - NOVEMBRE TEST.txt"), bank1Lines);

            File tempBank2 = new File(tempBank.getAbsoluteFile() + File.separator + accountBank2 + File.separator);
            tempBank2.mkdir();
            List<String> bank2Lines1 = Arrays.asList(
                    "BANQUE                   RELEVÉ DE COMPTE (EN EUR)",
                    "                         Date d'arrêté : 30/08/2024",
                    "Numéro de compte : ",
                    "IBAN :  BIC :  ",
                    "Date opé   Date valeur Libellé des opérations         Débit   Crédit ",
                    "                      L Ancien solde au 31/07/2024               0,00",
                    "17/08/2024 18/08/2024 REM CHQ N° 000000001                     100,00",
                    "Avec un text ",
                    "sur plusieur ligne 3T2024",
                    "                      Total des opérations             0,00    100,00",
                    "                      Nouveau solde au 30/05/2024              100,00"
            );
            Files.write(tempTestDir.resolve(tempBank2.getAbsoluteFile() + File.separator + accountBank2 + " -  AOUT TEST.txt"), bank2Lines1);

            List<String> bank2Lines2 = Arrays.asList(
                    "BANQUE                   RELEVÉ DE COMPTE (EN EUR)",
                    "                         Date d'arrêté : 31/07/2024",
                    "Numéro de compte : ",
                    "IBAN :  BIC :  ",
                    "Date opé   Date valeur Libellé des opérations         Débit   Crédit ",
                    "                      L'ancien solde au 30/06/2024               0,00",
                    "15/07/2024 16/07/2024 REM CHQ N° 000000001                     100,00",
                    "                      Total des opérations             0,00    100,00",
                    "                      Nouveau solde au 31/07/2024              100,00"
            );
            Files.write(tempTestDir.resolve(tempBank2.getAbsoluteFile() + File.separator + accountBank2 + " - JUILLET TEST.txt"), bank2Lines2);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public String createExpenseListFile() {
        try {
            List<String> lines = Arrays.asList(
                    "C'est le nom du Syndic",
                    "8 AVENUE DES CHAMPS ELYSE 11/04/2025 Page : 1",
                    "",
                    "75000 PARIS Liste des dépenses arrêtée au 31/12/2024",
                    "",
                    "Pièce Date              Libellé                                     Montant      Déduction    Récuperation",
                    "",
                    "",
                    "Clé : 001 C'est la première clé",
                    "",
                    "Code Nature : 61100 Nettoyage",
                    "33333 01/01/2024     Facture de nettoyage                           100.00 €     0.00 €       0.00 €",
                    "44444 15/02/2024     Facture de nettoyage supplémentaire            150.00 €     0.00 €       0.00 €",
                    "Total Nature : 61100                                                250.00 €     0.00 €       0.00 €",
                    "",
                    "Code Nature : 61500 Entretien",
                    "55555 10/03/2024     Facture d'entretien                            200.00 €     0.00 € ",
                    "Total Nature : 61500                                                200.00 €     0.00 €       0.00 €",
                    "",
                    "Total clé :    001                                                  450.00 €     0.00 €       0.00 €",
                    "",
                    "Total immeuble :                                                    570.00 €     0.00 €       0.00 €"
            );
            Files.write(tempTestDir.resolve("Liste_des_depenses.txt"), lines);
            return tempTestDir + File.separator + "Liste_des_depenses.txt";
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
