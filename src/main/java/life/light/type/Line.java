package life.light.type;

public record Line(String document, String date, TypeAccount account, String journal, TypeAccount accountCounterpart,
                   String checkNumber, String label, String debit, String credit) {
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
