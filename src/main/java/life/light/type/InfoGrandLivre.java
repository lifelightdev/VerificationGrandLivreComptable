package life.light.type;

import java.time.LocalDate;

public record InfoGrandLivre(String syndicName, LocalDate printDate, LocalDate stopDate, String postalCode) {
}
