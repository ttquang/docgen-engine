package com.example.docx4jexample1;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class DocGenController {

    @Autowired
    private DocGenService docGenService;

    @PostMapping(value = "/execute", produces = MediaType.APPLICATION_XML_VALUE)
    public DocGenContext execute(@RequestBody DocGenContext context) {
        docGenService.execute(context);
        return context;
    }


}
