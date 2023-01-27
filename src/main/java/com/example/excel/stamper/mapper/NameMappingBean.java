package com.example.excel.stamper.mapper;

import java.util.List;
import java.util.ArrayList;

public class NameMappingBean  {

    private String name; 
    private String type;
    private List<String> values = new ArrayList<String>();

    public NameMappingBean() {
    }

    public NameMappingBean(String name) {
        this.name = name;
    }
  
    public NameMappingBean(String name, List<String> values) {
        this.name = name;
        this.values = values;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public List<String> getValues() {
        return values;
    }

    public void setValues(List<String> values) {
        this.values = values;
    }

    public void setValue(String value) {
        this.values.add(value);
    }
    
    public void setValue(double value) {
        this.values.add(String.valueOf(value));
    }

    public String toString() {
        return getName() + ":" + getValues().toString();
    }
    
}
