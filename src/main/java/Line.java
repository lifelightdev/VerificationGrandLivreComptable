public class Line {
    public String document;
    public String date;
    public Account account;
    public String journal;
    public String counterpart;
    public String label;
    public String debit;
    public String credit;

    public Line(String document, String date, Account account, String journal, String counterpart,
                String label, String debit, String credit) {
        this.document = document;
        this.date = date;
        this.account = account;
        this.journal = journal;
        this.counterpart = counterpart;
        this.label = label;
        this.debit = debit;
        this.credit = credit;
    }
}
