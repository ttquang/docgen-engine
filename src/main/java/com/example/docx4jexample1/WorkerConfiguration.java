package com.example.docx4jexample1;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.scheduling.TaskScheduler;
import org.springframework.scheduling.annotation.Async;
import org.springframework.scheduling.annotation.EnableAsync;
import org.springframework.scheduling.annotation.EnableScheduling;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.scheduling.concurrent.ThreadPoolTaskScheduler;

import java.util.Optional;
import java.util.Random;

@Configuration
@EnableScheduling
@EnableAsync
public class WorkerConfiguration {
    @Autowired
    private DocGenService docGenService;

//    @Bean
//    public ThreadPoolTaskScheduler threadPoolTaskScheduler() {
//        ThreadPoolTaskScheduler threadPoolTaskScheduler
//                = new ThreadPoolTaskScheduler();
//        threadPoolTaskScheduler.setPoolSize(5);
//        threadPoolTaskScheduler.setThreadNamePrefix("ThreadPoolTaskScheduler");
//        return threadPoolTaskScheduler;
//    }

    @Scheduled(fixedRate = 500)
    @Async
    public void scheduleFixedRateTaskAsync() throws InterruptedException {
//        Random random = new Random(10);
//        System.out.println(Thread.currentThread().getName() +
//                " Fixed rate task async - " + System.currentTimeMillis() / 1000);
//        Thread.sleep(random.nextInt(10) * 1000);
//        docGenService.addWork();
        Optional<OBDocGenQueue> itemOptional = docGenService.getWork();
        if (itemOptional.isPresent()) {
            OBDocGenQueue item = itemOptional.get();
            System.out.println(Thread.currentThread().getName() + " " + item.getId());
            DocGenContext context = new DocGenContext();
            context.setOracleXML("ORACLE_XML_TEST");
            context.setTemplate("CC_Application_Form_new");
            context.setApplicationNo("030CC63000105");
            docGenService.execute(context);
            item.setStatus("DONE");
            docGenService.completeWork(item);

        } else {
            docGenService.pickWork();
        }
    }

    @Bean
    public ThreadPoolTaskScheduler threadPoolTaskScheduler() {
        ThreadPoolTaskScheduler threadPoolTaskScheduler = new ThreadPoolTaskScheduler();
        threadPoolTaskScheduler.setPoolSize(10);
        threadPoolTaskScheduler.setThreadNamePrefix("DocGen-Worker-");
        return threadPoolTaskScheduler;
    }

}
