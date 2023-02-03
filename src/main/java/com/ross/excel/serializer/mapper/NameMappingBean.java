package com.ross.excel.serializer.mapper;

import java.util.List;
import java.util.ArrayList;

public class NameMappingBean  {

    private String name; 
    private List<Object> values = new ArrayList<Object>();

    public NameMappingBean() {
    }

    public NameMappingBean(String name) {
        this.name = name;
    }
  
    public NameMappingBean(String name, List<? extends Object> vals) {
        this.name = name;
        setValues(vals);
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<Object> getValues() {
        return values;
    }

    private void setValues(List<? extends Object> vals) {

        for (Object o : vals) { 
            if (o instanceof String)
                this.add((String) o);
            else if (o instanceof Double)
                this.add((Double) o);  
            else if (o instanceof Integer)
                this.add((Integer) o);    
        }
    }

    public void add(String value) {
        this.values.add(value);
    }
    
    public void add(Double value) {
        this.values.add(value);
    }

    public void add(Integer value) {
        this.values.add(value);
    }

    public String toString() {

        String s = getName() + ":" + getValues().toString();
        return (s.length() > 100 ? s.substring(0,100) + " ..." : s);
    }

}
