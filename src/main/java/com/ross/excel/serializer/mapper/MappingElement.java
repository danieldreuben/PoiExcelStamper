package com.ross.excel.serializer.mapper;

import java.util.List;
import java.util.ArrayList;


public class MappingElement {
    String name;
    String label; 
    Object value;

    public MappingElement(String name, String label, Object value) {
        this.name = name;
        this.label = label;
        this.value = value;
    }

    public static List<MappingElement> getInRows (List<NameMappingBean> beans) {
        List<MappingElement> beanrows = new ArrayList<MappingElement>();
        beans.forEach( (val) -> { 
            for (int i = 0; i < val.getValues().size(); i++)
                beanrows.add(new MappingElement(val.getName(), 
                    String.format("%s", i), val.getValues().get(i)));  
        });
        return beanrows;
    }
    public String toString() {
        return String.format("%s:%s:%s",name, label, value);
    }

    public String getLabel() {
        return label;
    }
}