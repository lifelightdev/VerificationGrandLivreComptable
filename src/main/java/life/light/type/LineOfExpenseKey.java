package life.light.type;

public record LineOfExpenseKey(String key, String label, String value, TypeOfExpense type) implements LineOfExpense {}
