package com.example.docx4jexample1;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.jpa.repository.query.Procedure;
import org.springframework.data.repository.query.Param;

import java.util.Optional;

public interface DocGenQueueRepository extends JpaRepository<OBDocGenQueue, Long> {

    @Query("select q from OBDocGenQueue q where q.worker = :worker and q.status = 'NEW'")
    Optional<OBDocGenQueue> findByWorker(@Param("worker") String worker);

    @Procedure(value = "DOCGEN_PICK_ITEM")
    void pickItem(String in_worker);
}
