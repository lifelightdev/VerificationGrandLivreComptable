public record Account(String account, String label) {
    @Override
    public String account() {
        return account;
    }

    @Override
    public String label() {
        return label;
    }
}
