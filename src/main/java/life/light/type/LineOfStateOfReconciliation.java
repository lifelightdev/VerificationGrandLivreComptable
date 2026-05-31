package life.light.type;

import java.time.LocalDate;

public record LineOfStateOfReconciliation(String ledgerDocument, String ledgerDate, TypeAccount ledgerAccount,
                                          String ledgerJournal, TypeAccount ledgerAccountCounterpart,
                                          String ledgerCheckNumber, String ledgerLabel, Double ledgerDebit,
                                          Double ledgerCredit, String bankMonth, LocalDate bankTransactionDate,
                                          LocalDate bankValueDate, String bankLabel, Double bankDebit,
                                          Double bankCredit,
                                          String bankComment) {
    public LineOfStateOfReconciliation(LineLedger lineLedger, BankLine bankLine, String comment) {
        this(lineLedger.document(), lineLedger.date(), lineLedger.account(), lineLedger.journal(),
                lineLedger.accountCounterpart(), lineLedger.checkNumber(), lineLedger.label(), lineLedger.amountDebit(),
                lineLedger.amountCredit(), bankLine.year() + "-" + bankLine.mounth(), bankLine.operationDate(),
                bankLine.valueDate(), bankLine.label(), bankLine.debit(), bankLine.credit(), comment);
    }
}
