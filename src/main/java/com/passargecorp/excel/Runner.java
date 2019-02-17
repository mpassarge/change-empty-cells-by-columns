package com.passargecorp.excel;

import com.passargecorp.excel.workbook.Work;
import com.passargecorp.excel.workbook.output.OutputInterface;
import com.passargecorp.excel.workbook.output.UniqueOutput;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Hello world!
 */
public class Runner {

    private static final Logger LOG = LoggerFactory.getLogger(Runner.class);

    public static void main(String[] args) {

        LOG.info("Starting with {} arguments", args);
        // TODO: 2/16/19 add Input from command line apache commons cli
        OutputInterface output = new UniqueOutput("S2 Lauren");

        Work.runner("T", "/home/mpassarge/Downloads/inter-rater-reliability-dev.xlsx", output);
        LOG.info("Finished.");
    }
}