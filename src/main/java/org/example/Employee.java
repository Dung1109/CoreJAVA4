package org.example;

import java.util.HashMap;
import java.util.Map;

public class Employee {

    private static int index = 0;
    private int EMP_INDEX;
    private String code;
    private String name;

    public Employee(String code, String name) {
        this.code = code;
        this.name = name;
    }

    public Employee() {
        EMP_INDEX = index;
        index += 1;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getCode() {
        return code;
    }

    public String getName() {
        return name;
    }



}
