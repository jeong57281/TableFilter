package com.test.tablefilter;

import android.Manifest;
import android.content.Intent;
import android.os.Bundle;
import android.view.View;
import android.widget.AdapterView;
import android.widget.Button;
import android.widget.ListView;

import androidx.appcompat.app.AppCompatActivity;

import java.util.List;

public class MainActivity extends AppCompatActivity {

    ListView recent_list;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        /* Storage READ 권한 획득
        api 23 이하에서는 항상 권한이 부여되므로 필요가 없다.
        s4 에서 테스트할 경우에는 주석처리
        requestPermissions(new String[]
                        {Manifest.permission.READ_EXTERNAL_STORAGE,
                         Manifest.permission.WRITE_EXTERNAL_STORAGE},
                2);
         */

        /*
        Button search_btn = (Button) findViewById(R.id.search_button);
        search_btn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Intent search_intent_act = new Intent(getApplicationContext(), searchActivity.class);
                startActivity(search_intent_act);
            }
        });
        */

        recent_list = (ListView) findViewById(R.id.recent_list);
        recent_list.setOnItemClickListener(new AdapterView.OnItemClickListener() {
            @Override
            public void onItemClick(AdapterView<?> parent, View view, int position, long id) {
                /* 최근 열어본 xls 파일의 경로와 정보가 함께 저장되어 리스트로 출력해야 함 */
            }
        });

        Button file_btn = (Button) findViewById(R.id.file_button);
        file_btn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Intent file_intent_act = new Intent(getApplicationContext(), AndroidExplorerActivity.class);
                startActivity(file_intent_act);
            }
        });
    }
}
