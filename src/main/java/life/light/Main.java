package life.light;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.*;
import java.time.LocalDateTime;
import java.time.temporal.ChronoUnit;
import java.util.HashMap;
import java.util.Map;

import static life.light.WriteFile.*;

public class Main {

    private static final Logger LOGGER = LogManager.getLogger();

    public static void main(String[] args) {
        LocalDateTime debut = LocalDateTime.now();
        LOGGER.info("Début à {}:{}:{}", debut.getHour(), debut.getMinute(), debut.getSecond());
        // Lire le fichier texte
        String fileName = ".\\temp\\fichier_fusionner.txt";
        String syndicName = "";
        String printDate = "";
        String stopDate = "";
        String postalCode = "";
        // Récuperation des informations pour la génération du nom de fichier
        try (BufferedReader reader = new BufferedReader(new FileReader(fileName))) {
            String line = reader.readLine();
            syndicName = ExtractInfo.syndicName(line);
            line = reader.readLine();
            printDate = ExtractInfo.printDate(line);
            reader.readLine();
            line = reader.readLine();
            stopDate = ExtractInfo.stopDate(line);
            postalCode = ExtractInfo.postalCode(line);
        } catch (IOException e) {
            LOGGER.error("Erreur lors de la lecture du fichier avec cette erreur {}", e.getMessage());
        }
        LOGGER.info("Le nom du syndic est : {}", syndicName);
        LOGGER.info("La date d'édition est le {}", printDate);
        LOGGER.info("La date d'arrêt des comptes est le {}", stopDate);
        LOGGER.info("Le code postal du syndic est {}", postalCode);

        // Génération de la liste des comptes
        Map<String, Account> accounts = new HashMap<>();
        int numberOfLineInFile = 0;
        try (BufferedReader reader = new BufferedReader(new FileReader(fileName))) {
            String line;
            while ((line = reader.readLine()) != null) {
                if (ExtractInfo.isAcccount(line, postalCode, "018")) {
                    Account account = ExtractInfo.account(line);
                    accounts.put(account.account(), account);
                }
                if (!line.isEmpty()) {
                    numberOfLineInFile++;
                }
            }
        } catch (IOException e) {
            LOGGER.error("Erreur lors de la lecture du fichier avec cette erreur {}", e.getMessage());
        }
        LOGGER.info("Il y a {} comptes dans le grandlivre", accounts.size());

        writeFileCSVAccounts(accounts, "." + File.separator + "temp" + File.separator + "Plan comptable.csv");
        writeFileExcelAccounts(accounts, "." + File.separator + "temp" + File.separator + "Plan comptable.xlsx");

        // Géneration du grand livre
        Object[] grandLivres = new Object[numberOfLineInFile];
        int indexInGrandLivres = 0;
        try (BufferedReader reader = new BufferedReader(new FileReader(fileName))) {
            String line;
            while ((line = reader.readLine()) != null) {
                if (!line.isEmpty()) {
                    if (ExtractInfo.isLigne(line)) {
                        Line lineOfGrandLivre = ExtractInfo.line(line, accounts);
                        if (lineOfGrandLivre != null) {
                            grandLivres[indexInGrandLivres] = lineOfGrandLivre;
                            indexInGrandLivres++;
                        }
                    } else if (ExtractInfo.isTotalAccount(line)) {
                        TotalAccount totalAccount = ExtractInfo.totalAccount(line, accounts);
                        if (totalAccount != null) {
                            grandLivres[indexInGrandLivres] = totalAccount;
                            indexInGrandLivres++;
                        }
                    } else if (ExtractInfo.isTotalBuilding(line)){
                        TotalBuilding totalBuilding = ExtractInfo.totalBuilding(line);
                            grandLivres[indexInGrandLivres] = totalBuilding;
                            indexInGrandLivres++;
                    }
                }
            }
        } catch (IOException e) {
            LOGGER.error("Erreur lors de la lecture du fichier avec cette erreur {}", e.getMessage());
        }

        writeFileGrandLivre(grandLivres);
        String nameFile = printDate.substring(6) + "-" + printDate.substring(3, 5) + "-" + printDate.substring(0, 2)
                + " Grand livre " + syndicName.substring(0, syndicName.length() - 1).trim()
                + " au " + stopDate.substring(6) + "-" + stopDate.substring(3, 5) + "-" + stopDate.substring(0, 2)
                + ".xlsx";
        writeFileExcelGrandLivre(grandLivres, nameFile);

        // TODO Génération des journaux
        // TODO Écriture des journaux dans des fichiers Excel (un par journal)
        LocalDateTime fin = LocalDateTime.now();
        LOGGER.info("La durée du traitement est de {} secondes", ChronoUnit.SECONDS.between(debut, fin));
    }
}