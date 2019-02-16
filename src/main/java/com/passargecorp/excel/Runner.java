package com.passargecorp.excel;

import com.passargecorp.excel.workbook.TypeOfOutput;
import com.passargecorp.excel.workbook.Work;

/**
 * Hello world!
 *
 */
public class Runner
{
    public static void main( String[] args )
    {
        Work.runner("B","/home/mpassarge/Downloads/inter-rater-reliability-dev.xlsx", TypeOfOutput.STRING);
    }
}