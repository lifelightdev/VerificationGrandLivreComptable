package life.light;

public record TotalBuilding(String label, String debit, String credit) {
    @Override
    public String debit() {
        return debit;
    }

    @Override
    public String credit() {
        return credit;
    }

    public String label() {
        return label;
    }
}
