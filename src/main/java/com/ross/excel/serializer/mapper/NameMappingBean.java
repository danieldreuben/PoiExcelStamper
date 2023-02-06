package com.ross.excel.serializer.mapper;

import java.util.List;
import java.util.ArrayList;

import java.util.Date;

public class NameMappingBean  {

    private String name; 
    private List<Object> values = new ArrayList<Object>();
    private contentTypes conentType = contentTypes.MIXED; 

	public enum contentTypes {
		MIXED, NUMBER, DATE, STRING;

		public static contentTypes getRandom()  {
			contentTypes[] allopts = values();
			int rand = (int) (Math.random() * allopts.length);			
			return allopts[rand];
		}		
	};       

    public NameMappingBean() {
    }

    public NameMappingBean(String name) {

        this.name = name;
    }
  
    public NameMappingBean(String name, List<? extends Object> vals) {

        this.name = name;
        setValues(vals);
    }

    public contentTypes getContentType() {
        return this.conentType;
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
            add(o);
        }
    }

    public void add(Object o) {

        if (o instanceof String)
            this.add((String) o);
        else if (o instanceof Double)
            this.add((Double) o);  
        else if (o instanceof Integer)
            this.add((Integer) o);
        else if (o instanceof Date)
            this.add((Integer) o);            
    }   

    public void add(String value) {
        this.values.add(value);
        this.conentType = contentTypes.STRING;
    }
    
    public void add(Double value) {
        this.values.add(value);
        this.conentType = contentTypes.NUMBER;
    }

    public void add(Integer value) {
        this.values.add(value);
        this.conentType = contentTypes.NUMBER;
    }

    public void add(Date value) {
        this.values.add(value);
        this.conentType = contentTypes.DATE;
    }
    public String toString() {

        String s = getName() + ":" + getValues().toString();
        return (s.length() > 100 ? s.substring(0,100) + " ..." : s);
    }

}
