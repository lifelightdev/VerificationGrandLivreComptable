public class ExtractInfo {
    public static String syndicName(String ligne) {
        return ligne.replace("|","");
    }

    public static String printDate(String line) {
        String[] words = line.trim().split(" ");
        for (String word : words) {
            if (word.matches("^(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[0-2])/(19|20)\\d{2}$")){
                return word;
            }
        }
        return "";
    }

    public static String stopDate(String line) {
        return printDate(line);
    }

    public static Account account(String line) {
        String[] words = line.trim().split(" ");
        String account = words[0];
        StringBuilder label = new StringBuilder();
        for (String word : words) {
            if (!account.equals(word)){
                label.append(" ").append(word);
            }
        }
        return new Account( account, label.toString().trim());
    }
}
