package life.light.type;

import java.util.List;

public record StateOfReconciliation(List<LineOfStateOfReconciliation> find,
                                    List<LineOfStateOfReconciliation> noFindInLegder,
                                    List<LineOfStateOfReconciliation> noFindInBank) {
}
