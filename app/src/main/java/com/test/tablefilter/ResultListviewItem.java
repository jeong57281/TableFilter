package com.test.tablefilter;

public class ResultListviewItem {
    private String colItem;
    private String colValue;

    public ResultListviewItem(String colItem, String colValue){
        this.colItem = colItem;
        this.colValue = colValue;
    }

    public String getColItem(){
        return colItem;
    }

    public String getColValue(){
        return colValue;
    }
}
