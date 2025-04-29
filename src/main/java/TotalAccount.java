public record TotalAccount(String label, Account account, String debit, String credit) {

    @Override
    public String label() {
        return label;
    }

    @Override
    public Account account() {
        return account;
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
