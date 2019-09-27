package com.test.tablefilter;

import android.os.Bundle;
import android.view.View;
import android.widget.Button;
import android.widget.TextView;

import androidx.appcompat.app.AppCompatActivity;

import java.io.IOException;
import java.io.InputStream;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class MainActivity extends AppCompatActivity {

    // excel
    Workbook workbook = null;
    Sheet sheet = null;

    // view
    TextView editText;
    Button btnSearch;
    TextView tvStatus;
    TextView tvInsideMax, tvOutsideMax;
    TextView tvInsideMin, tvOutsideMin;

    // etc
    boolean flag = false; // 일치하는 항목이 없을 경우, 안내 메세지를 위한 flag

    // tmp
    TextView tvInsideName;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        btnSearch = (Button) findViewById(R.id.buttonSearch);
        editText = (TextView) findViewById(R.id.editText);
        tvStatus = (TextView) findViewById(R.id.tvStatus);

        tvInsideMax = (TextView) findViewById(R.id.tvInsideMax);
        tvOutsideMax = (TextView) findViewById(R.id.tvOutsideMax);

        tvInsideMin = (TextView) findViewById(R.id.tvInsideMin);
        tvOutsideMin = (TextView) findViewById(R.id.tvOutsideMin);

        btnSearch.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                Excel(editText.getText().toString());
            }
        });
    }

    public void Excel(String word){
        try {
            // 엑셀 파일 열기
            InputStream inputStream = getBaseContext().getResources().getAssets().open("MS20659K.xls");
            workbook = Workbook.getWorkbook(inputStream);

            // 엑셀 파일의 첫 번째 시트 인식
            sheet = workbook.getSheet(0);

            // 각 변수 선언 및 초기화
            int MaxColumn = 8;
            // 행
            int RowStart = 4;
            int RowEnd = sheet.getColumn(MaxColumn - 1).length;
            // 열
            int ColumnStart = 2;
            int ColumnEnd = sheet.getRow(0).length;

            flag = false;
            for (int row = RowStart; row < RowEnd; row++) {
                String excelload = sheet.getCell(ColumnStart, row).getContents();
                if (excelload.equals(word) || excelload.toLowerCase().equals(word)) {
                    tvInsideMax.setText(sheet.getCell(ColumnStart + 2, row).getContents());
                    tvOutsideMax.setText(sheet.getCell(ColumnStart + 2 + 2, row).getContents());
                    tvInsideMin.setText(sheet.getCell(ColumnStart + 2 + 1, row).getContents());
                    tvOutsideMin.setText(sheet.getCell(ColumnStart + 2 + 2 + 1, row).getContents());
                    tvStatus.setText("항목을 찾았습니다.");
                    flag = true;
                    //tvInsideName.setText(excelload);
                }
            }

            if(!flag){
                tvStatus.setText("일치하는 항목이 없습니다.");
                tvInsideMax.setText("value");
                tvOutsideMax.setText("value");
                tvInsideMin.setText("value");
                tvOutsideMin.setText("value");
            }

        } catch (IOException e){
            e.printStackTrace();
        } catch (BiffException e){
            e.printStackTrace();
        } finally {
            workbook.close();
        }
    }
}
