package life.light;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.time.format.DateTimeFormatter;

public class Constant {

    private static final Logger LOGGER = LogManager.getLogger();
    public static final String LECTURE_FICHIER = "Erreur lors de la lecture du fichier avec cette erreur {}";
    public static final String ERREUR_LORS_DE_LA_RECHERCHE_DU_COMPTE_SUR_LA_LIGNE = "Erreur lors de la recherche du compte {} sur la ligne {}";
    public static final String ERREUR_IL_MANQUE_DES_MONTANTS_SUR_LA_LIGNE_DE_TOTAL = "Erreur, il manque des montants sur la ligne de total {}";
    public static final String ERREUR_LORS_DE_LA_RECHERCHE_DU_COMPTE_SUR_LE_TOTAL_DU_COMPTE_DE_LA_LIGNE = "Erreur lors de la recherche du compte sur le total du compte {} de la ligne {}";
    public static final String IL_MANQUE_X_DANS_X_A_LA_LIGNE_X = "Il manque {} dans {} à la ligne {}";
    public static final String ERREUR_LORS_DE_L_ECRITURE_DANS_LE_FICHIER_DE_SORTIE = "Erreur lors de l'écriture dans le fichier de sortie '{}': {}";
    public static final String IL_N_Y_A_PAS_LE_MEME_NOMBRE_D_OPERATION_POINTEES_OK_DANS_LE_GRAND_LIVRE_ET_SUR_LES_RELEVES_DE_COMPTE = "Il n'y a pas le même nombre d'opération pointées OK dans le grand livre et sur les relevés de compte";

    public static final char EURO = '€';
    public static final int CENTURY = 2000;
    public static final String CSV = ".csv";
    public static final String XLSX = ".xlsx";
    public static final String REPORT_DE = "Report de";
    public static final String DIRECTORY_NAME_RESULTAT = "resultat";
    public static final String ACCOUNTING_PLAN = "Plan comptable ";
    public static final String TOTAL_DES_OPERATIONS = "Total des opérations";
    public static final String TOTAL_IMMEUBLE = "Total immeuble";

    public static final DateTimeFormatter DATE_FORMATTER = DateTimeFormatter.ofPattern("dd/MM/yyyy");

    public void logError(String message, String parameter1) {
        LOGGER.error(message, parameter1);
    }

    public void logError(String message, String parameter1, String parameter2) {
        LOGGER.error(message, parameter1, parameter2);
    }

    public void logError(String message, String parameter1, String parameter2, String parameter3) {
        LOGGER.error(message, parameter1, parameter2, parameter3);
    }

    public void logError(String message) {
        LOGGER.error(message);
    }
}
