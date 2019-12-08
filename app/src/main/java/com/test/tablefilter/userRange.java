package com.test.tablefilter;

public class userRange {
    private int topLeftRow;
    private int topLeftColumn;
    private int bottomRightRow;
    private int bottomRightColumn;
    private String topLeftContents;

    public userRange(){
        this.topLeftContents = "test";
    }

    public void setTopLeftRow(int topLeftRow){
        this.topLeftRow = topLeftRow;
    }

    public void setTopLeftColumn(int topLeftColumn){
        this.topLeftColumn = topLeftColumn;
    }

    public void setBottomRightRow(int bottomRightRow){
        this.bottomRightRow = bottomRightRow;
    }

    public void setBottomRightColumn(int bottomRightColumn){
        this.bottomRightColumn = bottomRightColumn;
    }

    public void setTopLeftContents(String topLeftContents){
        this.topLeftContents = topLeftContents;
    }

    public int getTopLeftRow(){
        return topLeftRow;
    }

    public int getTopLeftColumn(){
        return topLeftColumn;
    }

    public int getBottomRightRow(){
        return bottomRightRow;
    }

    public int getBottomRightColumn(){
        return bottomRightColumn;
    }

    public String getTopLeftContents(){
        return topLeftContents;
    }
}
