package com.example.docx4jexample1;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.jdbc.core.namedparam.MapSqlParameterSource;
import org.springframework.jdbc.core.namedparam.SqlParameterSource;
import org.springframework.jdbc.core.simple.SimpleJdbcCall;
import org.springframework.stereotype.Component;

import java.util.List;
import java.util.Map;

@Component
public class StoredProcedure1 {
    @Autowired
    private JdbcTemplate jdbcTemplate;

//    private SimpleJdbcCall simpleJdbcCall;

    public String generateData(DocGenContext context) {
//        jdbcTemplate.setResultsMapCaseInsensitive(true);

        SimpleJdbcCall simpleJdbcCall = new SimpleJdbcCall(jdbcTemplate).withProcedureName(context.getOracleXML());
//        SimpleJdbcCall simpleJdbcCall = new SimpleJdbcCall(jdbcTemplate).withProcedureName("ORACLE_XML_TEST");

        SqlParameterSource in = new MapSqlParameterSource()
                .addValue("in_application_no", context.getApplicationNo());

        String result = "";

        try {
            Map out = simpleJdbcCall.execute(in);
            List<Map<String, String>> xmlOutputs = (List<Map<String, String>>) out.get("OUTPUT");
            result = xmlOutputs.get(0).get("XML_STR");
        } catch (Exception e) {
            e.printStackTrace();
        }

        return result;
    }
}
