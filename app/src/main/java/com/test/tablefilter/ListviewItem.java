package com.test.tablefilter;

public class ListviewItem {
    private int icon;
    private String name;

    public ListviewItem(int icon, String name){
        this.icon = icon;
        this.name = name;
    }

    public int getIcon(){
        return icon;
    }

    public String getName(){
        return name;
    }
}
