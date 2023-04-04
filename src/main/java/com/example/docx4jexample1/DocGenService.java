package com.example.docx4jexample1;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.util.Optional;
import java.util.UUID;

@Service
public class DocGenService {
    @Autowired
    private TemplateContainer templateContainer;

    @Autowired
    private StoredProcedure1 storedProcedure1;

    @Autowired
    private DocGenQueueRepository docGenQueueRepository;

    public String execute(DocGenContext context) {
        try {
            String data = storedProcedure1.generateData(context);
            DataSource dataSource = DataSourceFactory.constructXMLDataSource(new ByteArrayInputStream(data.getBytes()));
            WordMLPackageWrapper wordMLPackageWrapper =
                    new WordMLPackageWrapper(templateContainer.get(context.getTemplate() + ".docx"), dataSource);
            wordMLPackageWrapper.process();
            wordMLPackageWrapper.save(new File( "C:\\Users\\Admin\\Desktop\\Working\\docgen-output\\"+ context.getTemplate() + UUID.randomUUID() + ".docx"));
        } catch (Exception e) {
            e.printStackTrace();
        }

        return "";
    }

    public void addWork() {
        OBDocGenQueue item = new OBDocGenQueue();
        item.setName(UUID.randomUUID().toString());
        item.setWorker(Thread.currentThread().getName());
        item.setStatus("NEW");
        docGenQueueRepository.save(item);
    }

    public Optional<OBDocGenQueue> getWork() {
        return docGenQueueRepository.findByWorker(Thread.currentThread().getName());
    }
    public void pickWork() {
        docGenQueueRepository.pickItem(Thread.currentThread().getName());
    }

    public void completeWork(OBDocGenQueue item) {
        docGenQueueRepository.save(item);
    }

}
