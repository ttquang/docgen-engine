package com.example.docx4jexample1;

import org.apache.commons.jexl3.*;

import java.util.Map;
import java.util.Objects;

public class ConditionEvaluationUtils {
    private static JexlEngine jexl = new JexlBuilder().create();

    public static boolean evaluate(String expressionCondition, Map<String, Object> inputs) {
        boolean result = false;
        if (expressionCondition.isBlank()) {
            try {
                JexlExpression e = jexl.createExpression(expressionCondition);

                JexlContext jc = new MapContext();
                for (String input : inputs.keySet()) {
                    jc.set(input, inputs.get(input));
                }

                Object obj = e.evaluate(jc);
                if (Objects.nonNull(obj)) {
                    result = (boolean) obj;
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return result;
    }

}
