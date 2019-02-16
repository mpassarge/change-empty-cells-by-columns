package com.passargecorp.excel.workbook;

public enum TypeOfOutput {

    STRING{
        public String getValue() {
            return null;
        }
    }, UNIQUE {
        public String getValue() {
            return null;
        }
    }, ZERO {
        public String getValue() {
            return null;
        }
    };
    abstract public String getValue();
}
