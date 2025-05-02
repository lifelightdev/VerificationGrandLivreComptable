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
        System.out.println("Début à " + debut.getHour() + ":" + debut.getMinute() + ":" + debut.getSecond());
        // Lire le fichier texte
        File dossierEnLigne = new File(".\\temp\\TXT\\EnModeLigne");
        File[] fichiersEnLigne = dossierEnLigne.listFiles();
        String syndicName = "";
        String printDate = "";
        String stopDate = "";
        String postalCode = "";
        // Récuperation des informations pour la génération du nom de fichier
        try (BufferedReader reader = new BufferedReader(new FileReader(fichiersEnLigne[0]))) {
            String line = reader.readLine();
            syndicName = ExtractInfo.syndicName(line);
            line = reader.readLine();
            printDate = ExtractInfo.printDate(line);
            line = reader.readLine();
            line = reader.readLine();
            postalCode = ExtractInfo.postalCode(line);
            stopDate = ExtractInfo.stopDate(line);
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
        if (fichiersEnLigne != null) {
            for (File fichier : fichiersEnLigne) {
                if (fichier.getName().contains("-EnModeLigne")) {
                    try (BufferedReader reader = new BufferedReader(new FileReader(fichier.getAbsoluteFile()))) {
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
                        LOGGER.error("Erreur lors de la lecture du fichier {} avec cette erreur {}", fichier.getName(), e.getMessage());
                    }
                }
            }
        }
        LOGGER.info("Il y a {} comptes dans le grandlivre", accounts.size());

        //writeFileAccounts(accounts);
        //writeFileExcelAccounts(accounts);

        // Géneration du grand livre
        Object[] grandLivres = new Object[numberOfLineInFile];
        int indexInGrandLivres = 0;
        int numeroDePage = 1;
        File dossierEnColonne = new File(".\\temp\\TXT\\EnModeColonne");
        File[] fichiersEnColonne = dossierEnColonne.listFiles();
        if (fichiersEnLigne != null) {
            for (File fichierLigne : fichiersEnLigne) {
                if (fichierLigne.getName().contains(addZeros(numeroDePage))) {
                    if (fichierLigne.getName().contains("-EnModeLigne")) {
                        try (BufferedReader reader = new BufferedReader(new FileReader(fichierLigne.getAbsoluteFile()))) {
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
                                    }
                                }
                            }
                        } catch (IOException e) {
                            LOGGER.error("Erreur lors de la lecture du fichier {} avec cette erreur {}", fichierLigne.getName(), e.getMessage());
                        }
                    }
                }

                File fichierColonne = fichiersEnColonne[numeroDePage - 1];
                if (fichierColonne.getName().contains(addZeros(numeroDePage))) {
                    if (fichierColonne.getName().contains("-EnModeColonne")) {
                        String typpeColonne = "";
                        String colonneDebit = "";
                        String ligneDebit = "";
                        String colonneCredit = "";
                        String ligneCredit = "";
                        try (BufferedReader reader = new BufferedReader(new FileReader(fichierColonne.getAbsoluteFile()))) {
                            String line;
                            while ((line = reader.readLine()) != null) {
                                if ("Page :".equals(line)) {
                                    // C'est la colonne débit
                                    typpeColonne = "Debit";
                                    line = reader.readLine();
                                    indexInGrandLivres = 0;
                                    colonneDebit = "";
                                }
                                if ("C peb] — créan".equals(line)) {
                                    // C'est la colonne crédit
                                    typpeColonne = "Credit";
                                    line = reader.readLine();
                                    indexInGrandLivres = 0;
                                    colonneCredit = "";
                                }
                                if (grandLivres[indexInGrandLivres] instanceof Line) {
                                    if ("Debit".equals(typpeColonne)) {
                                        colonneDebit += "{" + line + "}";
                                        Line lineOfGrandLivre = (Line) grandLivres[indexInGrandLivres];
                                        ligneDebit += "{" + lineOfGrandLivre.debit() + "}";
                                        if (!lineOfGrandLivre.debit().equals(line)) {
                                            Line lineDebit = new Line(lineOfGrandLivre.document(),
                                                    lineOfGrandLivre.date(), lineOfGrandLivre.account(),
                                                    lineOfGrandLivre.journal(), lineOfGrandLivre.counterpart(),
                                                    lineOfGrandLivre.checkNumber(), lineOfGrandLivre.label(),
                                                    line, lineOfGrandLivre.credit());
                                            grandLivres[indexInGrandLivres] = lineDebit;
                                            indexInGrandLivres++;
                                        }
                                    }
                                }
                                if (grandLivres[indexInGrandLivres] instanceof Line) {
                                    if ("Credit".equals(typpeColonne)) {
                                        colonneCredit += "{" + line.replace("\\","").replace("%","") + "}";
                                        Line lineOfGrandLivre = (Line) grandLivres[indexInGrandLivres];
                                        ligneCredit += "{" + lineOfGrandLivre.credit() + "}";
                                        if (!lineOfGrandLivre.credit().equals(line)) {
                                            Line lineCredit = new Line(lineOfGrandLivre.document(),
                                                    lineOfGrandLivre.date(), lineOfGrandLivre.account(),
                                                    lineOfGrandLivre.journal(), lineOfGrandLivre.counterpart(),
                                                    lineOfGrandLivre.checkNumber(), lineOfGrandLivre.label(),
                                                    lineOfGrandLivre.debit(), line);
                                            grandLivres[indexInGrandLivres] = lineCredit;
                                            indexInGrandLivres++;
                                        }
                                    }
                                }
                            }
                            LOGGER.info("Colonne débit = {}", colonneDebit);
                            LOGGER.info("Ligne débit = {}", ligneDebit);
                            LOGGER.info("Colonne crédit = {}", colonneCredit);
                            LOGGER.info("Ligne crédit = {}", ligneCredit);
                        } catch (IOException e) {
                            LOGGER.error("Erreur lors de la lecture du fichier {} avec cette erreur {}", fichierColonne.getName(), e.getMessage());
                        }
                    }
                }
                numeroDePage++;
            }
        }

        writeFileGrandLivre(grandLivres);
        String nameFile = printDate.substring(6) + "-" + printDate.substring(3, 5) + "-" + printDate.substring(0, 2)
                + " Grand livre " + syndicName.substring(0, syndicName.length() - 1).trim()
                + " au " + stopDate.substring(6) + "-" + stopDate.substring(3, 5) + "-" + stopDate.substring(0, 2)
                + ".xlsx";
        //writeFileExcelGrandLivre(grandLivres, nameFile);

        // TODO Génération des journaux
        // TODO Écriture des journaux dans des fichiers Excel (un par journal)
        LocalDateTime fin = LocalDateTime.now();
        System.out.println("Début à " + debut.getHour() + ":" + debut.getMinute() + ":" + debut.getSecond());
        System.out.println("Fin à " + fin.getHour() + ":" + fin.getMinute() + ":" + fin.getSecond());
        System.out.println("La durée du traitement est de " + ChronoUnit.SECONDS.between(debut, fin) + " secondes");
    }

    private static String addZeros(int number) {
        int length = 3;
        StringBuilder numberString = new StringBuilder(String.valueOf(number));
        while (numberString.length() < length) {
            numberString.insert(0, "0");
        }
        return numberString.toString();
    }
}