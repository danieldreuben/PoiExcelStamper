package com.ross.excel.serializer.mapper;

import java.util.List;
import java.util.Arrays;
import java.util.ArrayList;
import java.util.Date;

public class NameMappingBean  {

    private String name; 
    private List<Object> values = new ArrayList<Object>();

    // used to gengerate test mapping beans but could be used to query or 
    // label the beans content 
    // 
	public enum contentTypes {
		EMPTY, MIXED, NUMBER, DATE, STRING;

		public static contentTypes getRandom()  {
			contentTypes[] allopts = values();
			int rand = (int) (Math.random() * allopts.length);			
			return allopts[rand];
		}		

        public static contentTypes getRandom(List<contentTypes> inclusions)  {
			contentTypes[] allopts = values();
            contentTypes n = null;
            //List<contentTypes> excludes = Arrays.asList(MIXED, EMPTY);
            do {           
			    final int rand = (int) (Math.random() * allopts.length);
                n = inclusions.stream()
                .filter(ex -> ex.equals(allopts[rand]))
                .findFirst()
                .orElse(null);    

            } while ( n == null ); 
			return n;
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

    // not implemented but would iterate over valeus and return 
    // beans content type

    public contentTypes getContentType() {
        return null;
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
    }
    
    public void add(Double value) {
        this.values.add(value);
    }

    public void add(Integer value) {
        this.values.add(value);
    }

    public void add(Date value) {
        this.values.add(value);
    }
    public String toString() {

        String s = getName() + ":" + getValues().toString();
        return (s.length() > 100 ? s.substring(0,100) + " ..." : s);
    }

}
