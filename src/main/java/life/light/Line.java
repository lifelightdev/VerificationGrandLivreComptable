package life.light;

public record Line (String document, String date, Account account, String journal, Account accountCounterpart,
                    String checkNumber, String label, String debit, String credit) {
    @Override
    public String document() {
        return document;
    }

    @Override
    public String date() {
        return date;
    }

    @Override
    public Account account() {
        return account;
    }

    @Override
    public String journal() {
        return journal;
    }

    @Override
    public Account accountCounterpart() {
        return accountCounterpart;
    }

    @Override
    public String checkNumber() {
        return checkNumber;
    }

    @Override
    public String label() {
        return label;
    }

    @Override
    public String debit() {
        return debit;
    }

    @Override
    public String credit() {
        return credit;
    }
}
