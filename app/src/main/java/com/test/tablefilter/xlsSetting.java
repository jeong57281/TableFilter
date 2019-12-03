package com.test.tablefilter;

import android.content.Intent;
import android.graphics.Bitmap;
import android.graphics.Canvas;
import android.graphics.Color;
import android.graphics.Paint;
import android.os.Bundle;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.ImageView;
import android.widget.TextView;

import androidx.annotation.Nullable;
import androidx.appcompat.app.AppCompatActivity;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class xlsSetting extends AppCompatActivity {

    Button btn;
    Button down_btn;
    Button up_btn;
    Button right_btn;
    Button left_btn;
    Button first_btn;
    Button last_btn;
    TextView tv_fileName;

    // excel
    Workbook workbook = null;
    Sheet sheet = null;

    // canvas
    Bitmap bitmap;
    Paint paint;
    ImageView imageView;
    Canvas canvas;

    int textsize = 30;

    /* 셀의 크기 지정 */
    int index_size = 40;
    int row_size = 160;
    int col_size = 40;

    /* 셀의 개수 지정 */
    int row_end = 4;
    int col_end = 15;

    /* 미리보기 범위를 조정할 크기값 */
    int row_move_size = 0;
    int col_move_size = 0;

    /* sheet 의 최대 행, 열 값 */
    int max_rows;
    int max_cols;

    @Override
    protected void onCreate(@Nullable Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.xls_setting);

        Preview_Execl("none");

        btn = (Button) findViewById(R.id.search_btn);
        btn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Intent intent = new Intent(getApplicationContext(), searchActivity.class);
                startActivity(intent);
            }
        });

        down_btn = (Button) findViewById(R.id.down_btn);
        down_btn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Preview_Execl("down");
            }
        });

        up_btn = (Button) findViewById(R.id.up_btn);
        up_btn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Preview_Execl("up");
            }
        });

        right_btn = (Button) findViewById(R.id.right_btn);
        right_btn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Preview_Execl("right");
            }
        });

        left_btn = (Button) findViewById(R.id.left_btn);
        left_btn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Preview_Execl("left");
            }
        });

        first_btn = (Button) findViewById(R.id.first_btn);
        first_btn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Preview_Execl("first");
            }
        });

        last_btn = (Button) findViewById(R.id.last_btn);
        last_btn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Preview_Execl("last");
            }
        });
    }

    public void Preview_Execl(String direction){
        // bitmap 크기 설정, Canvas 에 연결 ---------------------
        bitmap = Bitmap.createBitmap(row_size * row_end + index_size,
                                    col_size * col_end + index_size,
                                            Bitmap.Config.ARGB_8888);
        canvas = new Canvas(bitmap);
        canvas.drawColor(Color.BLACK);

        // 비트맵을 출력할 ImageView 와 연결 ---------------------
        imageView = (ImageView) findViewById(R.id.sample_image);
        imageView.setImageBitmap(bitmap);

        // canvas 에 그리기 위한 도구인 paint 객체 선언 -----------
        paint = new Paint();
        paint.setColor(Color.RED);

        /*
        표에서는 거리로 따지기 때문에 행과 열 개념이 반대
        좌표(인덱스)와 셀의 개수는 다르다!!!
                좌표      셀
            i : 행, 가로 열, 가로
            j : 열, 세로 행, 세로
        */
        try {
            Intent intent = getIntent();
            String filePath = intent.getExtras().getString("filePath");
            String fileName = intent.getExtras().getString("fileName");

            tv_fileName = (TextView) findViewById(R.id.tv_fileName);
            tv_fileName.setText("파일명 : " + fileName);

            File file = new File(filePath);

            // 엑셀 파일 열기 --------------------------------------
            //InputStream inputStream = getBaseContext().getResources().getAssets().open(fileName);
            workbook = Workbook.getWorkbook(file);

            // 엑셀 파일의 첫 번째 시트 인식 -------------------------
            sheet = workbook.getSheet(0);

            // 불러온 엑셀 데이터의 최대 행, 열 값 --------------------
            max_rows = sheet.getRows();
            max_cols = sheet.getColumns();

            // 사각형 스타일 설정 (테두리만 그리기) -------------------
            paint.setStyle(Paint.Style.STROKE);
            paint.setStrokeWidth(1);

            // Word Style ----------------------------------------
            paint.setColor(Color.WHITE);
            paint.setTextSize(textsize);
            paint.setTextAlign(Paint.Align.CENTER);

            /* ---------------------------------------------------
            max_row_index, max_col_index는 현재의 최대 index 값을 보여주므로
            엑셀 시트의 최대 데이터 범위보다 작을 경우에만 증가시켜야 한다.
            같을 경우에 중가시키면 시트의 범위를 벗어나게 된다.
            */

            int max_row_index = col_end + row_move_size;
            int max_col_index = row_end + col_move_size;

            switch(direction){
                case "down":
                    if(max_row_index < max_rows)
                        row_move_size++;
                    break;
                case "up":
                    if(row_move_size > 0)
                        row_move_size--;
                    break;
                case "right":
                    if(max_col_index < max_cols)
                        col_move_size++;
                    break;
                case "left":
                    if(col_move_size > 0)
                        col_move_size--;
                    break;
                case "last":
                    row_move_size = max_rows - col_end;
                    col_move_size = max_cols - row_end;
                    break;
                case "first":
                    row_move_size = 0;
                    col_move_size = 0;
                    break;
                default:
                    break;
            }

            // index 를 나타내는 0행 0열 ----------------------------
            paint.setColor(Color.GREEN);
            canvas.drawRect(0, 0, (row_size * row_end) + index_size, index_size, paint);
            canvas.drawRect(0, 0, index_size, (col_size * col_end) + index_size, paint);

            int k = 0;
            for(k = 0; k < row_end; k++){
                canvas.drawText(Integer.toString(k + 1 + col_move_size),
                        (k * row_size) + (row_size / 2) + index_size,
                        (0 * col_size) + (col_size*3/4),
                        paint);
            }

            for(k = 0; k < col_end; k++){
                canvas.drawText(Integer.toString(k + 1 + row_move_size),
                        (0 * row_size) + (index_size / 2),
                        (k * col_size) + (col_size*3/4) + index_size,
                        paint);
            }

            /* -----------------------------------------------------
            10x10 사이즈의 엑셀 cell 이 미리보기에 출력된다.
            따라서 10x10 보다 큰 엑셀 데이터만 사용하도록 권고.
            */
            Log.d("max_row_size", Integer.toString(max_rows));
            Log.d("col_end_size", Integer.toString(col_end));
            Log.d("row_move_size", Integer.toString(row_move_size));

            for (int i = 0; i < row_end; i++) {
                for (int j = 0; j < col_end; j++) {

                    String excelload = sheet.getCell(i + col_move_size, j + row_move_size).getContents();
                    if(excelload.equals("")){
                        paint.setColor(Color.WHITE);
                    }
                    else{
                        paint.setColor(Color.WHITE);
                        // end, start 값을 지정하면, 그만큼의 길이가 안되는 셀은 출력할때 오류 발생
                        if(excelload.length() > 5){
                            canvas.drawText(excelload,
                                    0,
                                    5,
                                    (i * row_size) + (row_size / 2) + index_size,
                                    (j * col_size) + (col_size*3/4) + index_size,
                                    paint);
                        }
                        else{
                            canvas.drawText(excelload,
                                    (i * row_size) + (row_size / 2) + index_size,
                                    (j * col_size) + (col_size*3/4) + index_size,
                                    paint);
                        }
                        paint.setColor(Color.BLUE);
                    }

                    /* Paint Cell */
                    canvas.drawRect((i * row_size) + index_size,
                            (j * col_size) + index_size,
                            (i * row_size) + row_size + index_size,
                            (j * col_size) + col_size + index_size,
                            paint);
                }
            }
        } catch (
                IOException e){
            e.printStackTrace();
        } catch (BiffException e){
            e.printStackTrace();
        } finally {
            workbook.close();
        }
    }
}
