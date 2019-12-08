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


    /*  1. count 즉, 셀의 갯수는 행, 열의 개념이 반대
        2. 병합된 topLeft, bottomRight Cell 은 병합된 쉘에 함께 포함되어 있기 때문에 +1 을 해주어야 함. */
    public int getMergeRowCount(){
        return getBottomRightColumn() - getTopLeftColumn() + 1;
    }

    public int getMergeColCount(){
        return getBottomRightRow() - getTopLeftRow() + 1;
    }
}
