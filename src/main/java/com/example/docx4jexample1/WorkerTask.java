package com.example.docx4jexample1;

import java.util.Date;

public class WorkerTask implements Runnable {
    private String message;

    public WorkerTask(String message) {
        this.message = message;
    }

    @Override
    public void run() {
        System.out.println(new Date() + " Runnable Task with " + message
                + " on thread " + Thread.currentThread().getName());
    }
}