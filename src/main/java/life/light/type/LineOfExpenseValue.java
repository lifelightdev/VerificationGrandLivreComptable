package life.light.type;

import java.time.LocalDate;

public record LineOfExpenseValue(String document, LocalDate date, String label, String amount, String deduction,
                                 String recovery, String natureCode, String keyCode) implements LineOfExpense {
}