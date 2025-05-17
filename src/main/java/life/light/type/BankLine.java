package life.light.type;

import java.time.LocalDate;

public record BankLine(Integer year, Integer mounth, LocalDate operationDate, LocalDate valueDate, TypeAccount account,
                       String label, Double debit, Double credit) {
}