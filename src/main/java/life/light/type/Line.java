package life.light.type;

public record Line(String document, String date, TypeAccount account, String journal, TypeAccount accountCounterpart,
                   String checkNumber, String label, String debit, String credit) {
}
