package life.light.type;

public record LineOfExpense(String document, java.time.LocalDate date, String label, String amount, String deduction,
                            String recovery) {
}