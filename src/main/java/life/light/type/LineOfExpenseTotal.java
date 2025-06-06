package life.light.type;

public record LineOfExpenseTotal(String key, String label, String value, String amount, String deduction,
                                 String recovery, TypeOfExpense type) implements LineOfExpense{
}
