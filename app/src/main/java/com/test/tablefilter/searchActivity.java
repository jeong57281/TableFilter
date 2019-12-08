package com.test.tablefilter;

import android.content.Intent;
import android.os.Bundle;
import android.util.Log;
import android.view.View;
import android.widget.ArrayAdapter;
import android.widget.AutoCompleteTextView;
import android.widget.Button;
import android.widget.ListView;
import android.widget.TextView;

import androidx.appcompat.app.AppCompatActivity;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import javax.xml.transform.Result;

import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;

public class searchActivity extends AppCompatActivity {
    // excel
    Workbook workbook = null;
    Sheet sheet = null;
    String[][] excelArray;

    // view
    Button btnSearch;
    TextView tvStatus;
    ListView result_ResultList;
    AutoCompleteTextView autoCompleteTextView;

    // etc
    boolean flag = false; // 일치하는 항목이 없을 경우, 안내 메세지를 위한 flag

    /* intent */
    Intent intent;

    // 전달받은 (표의)좌표값
    int row_start, col_start;
    int row_end, col_end;

    // 전달받은 파일 경로
    String filePath;

    // ResultList
    List<String> ResultList, SearchList;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.search_activity);

        btnSearch = (Button) findViewById(R.id.buttonSearch);
        //editText = (TextView) findViewById(R.id.editText);
        tvStatus = (TextView) findViewById(R.id.tvStatus);
        result_ResultList = (ListView) findViewById(R.id.result_list);
        autoCompleteTextView = (AutoCompleteTextView) findViewById(R.id.autoCompleteTextView);

        // 출력 데이터를 저장하게 되는 리스트
        ResultList = new ArrayList<>();
        // 검색어 데이터를 저장하는 리스트
        SearchList = new ArrayList<>();

        // 리스트뷰와 리스트를 연결하기 위해 사용되는 어댑터
        ArrayAdapter<String> ResultAdapter = new ArrayAdapter<>(this,
                android.R.layout.simple_list_item_1, ResultList);

        ArrayAdapter<String> SearchAdapter = new ArrayAdapter<>(this,
                android.R.layout.simple_list_item_1, SearchList);

        // excel 데이터 배열에 불러오기
        Excel();

        // 리스트뷰와 자동완성 텍스트 뷰의 어댑터를 지정해준다.
        result_ResultList.setAdapter(ResultAdapter);

        autoCompleteTextView.setAdapter(SearchAdapter);

        btnSearch.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                SearchExcel(autoCompleteTextView.getText().toString());
            }
        });
    }

    public void Excel(){
        try {
            intent = getIntent();

            // 파일 경로
            filePath = intent.getExtras().getString("filePath");

            // 엑셀파일의 테이블 좌표
            row_start = intent.getExtras().getInt("row_start");
            //row_end = intent.getExtras().getInt("row_end");
            col_start = intent.getExtras().getInt("col_start");
            //col_end = intent.getExtras().getInt("col_end");

            File file = new File(filePath);

            // 엑셀 파일 열기 --------------------------------------
            //InputStream inputStream = getBaseContext().getResources().getAssets().open(fileName);

            /* jxl encoding setting : utf-8 */
            WorkbookSettings ws = new WorkbookSettings();
            ws.setEncoding("Cp1252");
            workbook = Workbook.getWorkbook(file, ws);

            // 엑셀 파일의 첫 번째 시트 인식
            sheet = workbook.getSheet(0);

            // 엑셀 파일에 표만 존재하고 불순물이 없다는 전제하에 마지막(end)값은 받지 않아도 됨.
            // 값은 1부터 시작하므로 1을 빼줌.
            row_end = sheet.getRows();
            col_end = sheet.getColumns();

            int table_row_count = row_end - row_start;
            int table_col_count = col_end - col_start;

            // java 배열은 순서대로 행, 열
            excelArray = new String[table_row_count][table_col_count];

            for (int row = row_start; row < row_end; row++) {
                for (int col = col_start; col < col_end; col++) {
                    Log.d("table row", Integer.toString(row));
                    Log.d("table col", Integer.toString(col));
                    // getCell 은 열, 행 순 (좌표 개념)
                    String cell_value = sheet.getCell(col, row).getContents();
                    excelArray[row - row_start][col - col_start] = cell_value;
                    Log.d("sheet value", sheet.getCell(col, row).getContents());
                }
            }

            for (int row = row_start + 1; row < row_end; row++) {
                String cell_value = excelArray[row - row_start][0];
                SearchList.add(cell_value);
            }

        } catch (IOException e){
            e.printStackTrace();
        } catch (BiffException e){
            e.printStackTrace();
        } finally {
            workbook.close();
        }
    }

    public void SearchExcel(String word){
        ResultList.clear();
        flag = false;
        // row, col 의 시작(start) 좌표는 0, 0
        // 1씩 값을 더해 (1, 0), (0, 1) 좌표를 이동
        for (int row = row_start + 1; row < row_end; row++) {
            String rowItem = excelArray[row - row_start][0];
            // 대소문자 구분 없이 입력된 값을 받아 비교
            Log.d("rowItem", rowItem);
            if (rowItem.equals(word) || rowItem.toLowerCase().equals(word)) {
                for(int col = col_start + 1; col < col_end; col++){
                    String colItem = excelArray[0][col - col_start];
                    ResultList.add(colItem);
                    String colValue = excelArray[row - row_start][col - col_start];
                    ResultList.add(colValue);
                }
                tvStatus.setText("항목을 찾았습니다.");
                flag = true;
            }
        }
        if(!flag){
            tvStatus.setText("일치하는 항목이 없습니다.");
        }
    }
}
