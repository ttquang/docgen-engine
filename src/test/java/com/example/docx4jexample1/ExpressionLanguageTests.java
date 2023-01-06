package com.example.docx4jexample1;

import org.apache.commons.jexl3.*;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

@SpringBootTest
class ExpressionLanguageTests {
    private JexlEngine jexl = new JexlBuilder().create();

    @Test
    void contextLoads() {
    }

    @Test
    void compare_text_test_equal() {
        // Create an expression
        String jexlExp = "foo == 'Y'";
        JexlExpression e = jexl.createExpression( jexlExp );

        // Create a context and add data
        JexlContext jc = new MapContext();
        jc.set("foo", "Y");

        // Now evaluate the expression, getting the result
        Object o = e.evaluate(jc);

        System.out.println(o);
    }

    @Test
    void compare_text_number_equal() {
        // Create an expression
        String jexlExp = "foo == 100";
        JexlExpression e = jexl.createExpression( jexlExp );

        // Create a context and add data
        JexlContext jc = new MapContext();
//        jc.set("foo", "100");
        jc.set("foo", 100);

        // Now evaluate the expression, getting the result
        Object o = e.evaluate(jc);

        System.out.println(o);
    }

    @Test
    void compare_text_number_greater_than() {
        // Create an expression
        String jexlExp = "foo > 100";
        JexlExpression e = jexl.createExpression( jexlExp );

        // Create a context and add data
        JexlContext jc = new MapContext();
//        jc.set("foo", "100");
        jc.set("foo", 101);

        // Now evaluate the expression, getting the result
        Object o = e.evaluate(jc);

        System.out.println(o);
    }
}
