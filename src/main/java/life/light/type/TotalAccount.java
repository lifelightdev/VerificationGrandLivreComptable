package life.light.type;

public record TotalAccount(String label, TypeAccount account, String debit, String credit) {
    public Double amountDebit(){
        String d = debit.replace(" ","").replace(",",".").replace("€","").trim();
        try {
            return Double.parseDouble(d);
        } catch (NumberFormatException e) {
            return 0.0;
        }
    }
    public Double amountCredit(){
        String c = credit.replace(" ","").replace(",",".").replace("€","").trim();
        try {
            return Double.parseDouble(c);
        } catch (NumberFormatException e) {
            return 0.0;
        }
    }
}
