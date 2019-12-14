package com.test.tablefilter;

import android.content.Context;
import android.view.LayoutInflater;
import android.view.View;
import android.view.ViewGroup;
import android.widget.BaseAdapter;
import android.widget.Button;
import android.widget.TextView;

import java.util.ArrayList;

public class RecentListviewAdapter extends BaseAdapter implements View.OnClickListener {

    // 버튼 클릭 이벤트를 위한 Listener 인터페이스 정의.
    public interface ListBtnClickListener {
        void onListBtnClick(int position);
    }

    private LayoutInflater inflater;
    private ArrayList<RecentListviewItem> data;
    private int layout; // resource id 저장
    private ListBtnClickListener listBtnClickListener; // 생성자로부터 전달된 ListBtnClickListener  저장.

    public RecentListviewAdapter(Context context
            , int layout
            , ArrayList<RecentListviewItem> data
            , ListBtnClickListener clickListener){
        this.inflater = (LayoutInflater) context.getSystemService(Context.LAYOUT_INFLATER_SERVICE);
        this.layout = layout;
        this.data = data;
        this.listBtnClickListener = clickListener;
    }

    @Override
    public int getCount() {
        return data.size();
    }

    @Override
    public String getItem(int position) {
        return data.get(position).getRecentFileName();
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

        RecentListviewItem listviewitem = data.get(position);

        TextView recentFileName = (TextView) convertView.findViewById(R.id.tv_recentFileName);
        recentFileName.setText(listviewitem.getRecentFileName());

        Button recentDeleteBtn = (Button) convertView.findViewById(R.id.recentDeleteBtn);
        recentDeleteBtn.setTag(position);
        recentDeleteBtn.setOnClickListener(this);

        return convertView;
    }

    @Override
    public void onClick(View v) {
        /* 생성자로부터 this(mainActivity 의 onListBtnClick() 함수)를 전달받으므로
        this.listBtnClickListener 에는 mainActivity 의 onListBtnClick() 함수가 호출된다. */
        if (this.listBtnClickListener != null) {
            this.listBtnClickListener.onListBtnClick((int)v.getTag()) ;
        }
    }
}
