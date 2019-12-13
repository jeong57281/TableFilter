package com.test.tablefilter;

import android.content.Context;
import android.view.LayoutInflater;
import android.view.View;
import android.view.ViewGroup;
import android.widget.BaseAdapter;
import android.widget.ImageView;
import android.widget.TextView;

import java.util.ArrayList;

public class ResultListviewAdapter extends BaseAdapter {

    private LayoutInflater inflater;
    private ArrayList<ResultListviewItem> data;
    private int layout;

    public ResultListviewAdapter(Context context, int layout, ArrayList<ResultListviewItem> data){
       this.inflater = (LayoutInflater) context.getSystemService(Context.LAYOUT_INFLATER_SERVICE);;
       this.layout = layout;
       this.data = data;
    }

    @Override
    public int getCount() {
        return data.size();
    }

    @Override
    public String getItem(int position) {
        return data.get(position).getColValue();
    }

    @Override
    public long getItemId(int position){
        return position;
    }

    @Override
    public View getView(int position, View convertView, ViewGroup parent){

        if(convertView == null){
            convertView = inflater.inflate(layout,parent,false);
        }

        ResultListviewItem listviewitem = data.get(position);

        TextView colItem = (TextView) convertView.findViewById(R.id.colItemRow);
        colItem.setText(listviewitem.getColItem());

        TextView colValue = (TextView) convertView.findViewById(R.id.colValueRow);
        colValue.setText(listviewitem.getColValue());

        return convertView;
    }
}
