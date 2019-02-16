package com.passargecorp.excel;

import com.passargecorp.excel.workbook.Work;
import com.passargecorp.excel.workbook.output.OutputInterface;
import com.passargecorp.excel.workbook.output.UniqueOutput;

/**
 * Hello world!
 */
public class Runner {
    public static void main(String[] args) {
        OutputInterface output = new UniqueOutput("S2 Lauren");
        Work.runner("T", "/home/mpassarge/Downloads/inter-rater-reliability-dev.xlsx", output);
    }
}