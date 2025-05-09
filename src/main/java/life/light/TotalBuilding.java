package life.light;

public record TotalBuilding(String debit, String credit) {
    @Override
    public String debit() {
        return debit;
    }

    @Override
    public String credit() {
        return credit;
    }
}
