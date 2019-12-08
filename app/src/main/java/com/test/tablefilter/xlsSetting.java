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
import android.widget.EditText;
import android.widget.ImageView;
import android.widget.TextView;

import androidx.annotation.Nullable;
import androidx.appcompat.app.AppCompatActivity;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

import jxl.Cell;
import jxl.Range;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
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

    // intent
    Intent intent, intent2;
    String fileName;
    String filePath;

    // excel
    Workbook workbook = null;
    Sheet sheet = null;
    String excelload, excelload_alpha;

    // canvas
    Bitmap bitmap;
    Paint paint;
    ImageView imageView;
    Canvas canvas;

    // 반복문 변수
    int i, j, k;

    int text_size = 30;

    /* .xls file 미리보기를 구현할 캔버스 정보 */
    // 셀의 크기 지정
    int row_index_size = 50;
    int col_index_size = 40;
    int row_size = 160;
    int col_size = 40;

    // 셀의 개수 지정
    int row_count = 15;
    int col_count = 4;

    //  범위를 조정할 크기값 */
    int row_move_count = 0;
    int col_move_count = 0;

    /* 실제 .xls sheet 의 최대 행, 열 개수 */
    int xls_row_max_count;
    int xls_col_max_count;

    /* 미리보기에서 이동한 행, 열의 현재 위치 */
    int current_row_count;
    int current_col_count;

    /* 입력받은 좌표값 */
    int row_start, row_end, col_start, col_end;

    EditText rowEdit_start, rowEdit_end, colEdit_start, colEdit_end;

    @Override
    protected void onCreate(@Nullable Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.xls_setting);

        /* intent */
        intent = getIntent();
        filePath = intent.getExtras().getString("filePath");
        fileName = intent.getExtras().getString("fileName");

        tv_fileName = (TextView) findViewById(R.id.tv_fileName);
        tv_fileName.setText("파일명 : " + fileName);

        /* 미리보기 첫 호출 */
        Preview_Execl("none");

        rowEdit_start = (EditText) findViewById(R.id.rowEdit_start);
        //rowEdit_end = (EditText) findViewById(R.id.rowEdit_end);
        colEdit_start = (EditText) findViewById(R.id.colEdit_start);
        //colEdit_end = (EditText) findViewById(R.id.colEdit_end);

        btn = (Button) findViewById(R.id.search_btn);
        btn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                intent2 = new Intent(getApplicationContext(), searchActivity.class);
                /* 파일 경로를 전달 */
                intent2.putExtra("filePath", filePath);

                /* 입력한 좌표값을 전달 */
                row_start = Integer.parseInt(rowEdit_start.getText().toString());
                //row_end = Integer.parseInt(rowEdit_end.getText().toString());
                col_start = Integer.parseInt(colEdit_start.getText().toString());
                //col_end = Integer.parseInt(colEdit_end.getText().toString());

                intent2.putExtra("row_start", row_start);
                //intent2.putExtra("row_end", row_end);
                intent2.putExtra("col_start", col_start);
                //intent2.putExtra("col_end", col_end);

                /* 만약 입력한 값이 없을 경우에 대한 처리 추가 필요 */

                startActivity(intent2);
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
        bitmap = Bitmap.createBitmap(row_size * col_count + row_index_size,
                                    col_size * row_count + col_index_size,
                                            Bitmap.Config.ARGB_8888);
        canvas = new Canvas(bitmap);
        canvas.drawColor(Color.BLACK);

        // 비트맵을 출력할 ImageView 와 연결 ---------------------
        imageView = (ImageView) findViewById(R.id.sample_image);
        imageView.setImageBitmap(bitmap);

        // canvas 에 그리기 위한 도구인 paint 객체 선언 -----------
        paint = new Paint();
        paint.setColor(Color.RED);

        // ----------------------------------------------------
        try {
            File file = new File(filePath);

            /* jxl encoding setting : utf-8 */
            WorkbookSettings ws = new WorkbookSettings();
            ws.setEncoding("Cp1252");

            // 엑셀 파일 열기 --------------------------------------
            //InputStream inputStream = getBaseContext().getResources().getAssets().open(fileName);
            workbook = Workbook.getWorkbook(file, ws);

            // 엑셀 파일의 첫 번째 시트 인식 -------------------------
            sheet = workbook.getSheet(0);

            // 불러온 엑셀 데이터의 최대 행, 열 값 --------------------
            xls_row_max_count = sheet.getRows();
            xls_col_max_count = sheet.getColumns();

            // 사각형 스타일 설정 (테두리만 그리기) -------------------
            paint.setStyle(Paint.Style.FILL);
            paint.setStrokeWidth(1);

            // Word Style ----------------------------------------
            paint.setColor(Color.BLACK);
            paint.setTextSize(text_size);
            paint.setTextAlign(Paint.Align.CENTER);

            /* ---------------------------------------------------
            current_row_count, current_col_count는 현재의 최대 index 값을 보여주므로
            엑셀 시트의 최대 데이터 범위보다 작을 경우에만 증가시켜야 한다.
            같을 경우에 증가시키면 시트의 범위를 벗어나게 된다.
            */

            current_row_count = row_count + row_move_count;
            current_col_count = col_count + col_move_count;

            switch(direction){
                case "down":
                    if(current_row_count < xls_row_max_count)
                        row_move_count++;
                    break;
                case "up":
                    if(row_move_count > 0)
                        row_move_count--;
                    break;
                case "right":
                    if(current_col_count < xls_col_max_count)
                        col_move_count++;
                    break;
                case "left":
                    if(col_move_count > 0)
                        col_move_count--;
                    break;
                case "last":
                    row_move_count = xls_row_max_count - row_count;
                    col_move_count = xls_col_max_count - col_count;
                    break;
                case "first":
                    row_move_count = 0;
                    col_move_count = 0;
                    break;
                default:
                    break;
            }

            /* -----------------------------------------------------
            4x15 사이즈의 엑셀 cell 이 미리보기에 출력된다.
            따라서 4x15 보다 큰 엑셀 데이터만 사용하도록 권고.
            Log.d("xls_row_max_count", Integer.toString(xls_row_max_count));
            Log.d("row_count", Integer.toString(row_count));
            Log.d("row_move_count", Integer.toString(row_move_count));
             */

            for (i = 0; i < col_count; i++) {
                for (j = 0; j < row_count; j++) {
                    excelload = sheet.getCell(i + col_move_count, j + row_move_count).getContents();
                    // 셀에 값이 있을 경우
                    if(!excelload.equals("")){
                        /* 배경 셀 그리기 */
                        paint.setStyle(Paint.Style.FILL);
                        paint.setColor(Color.WHITE);
                        canvas.drawRect((i * row_size) + row_index_size,
                                (j * col_size) + col_index_size,
                                (i * row_size) + row_size + row_index_size,
                                (j * col_size) + col_size + col_index_size,
                                paint);

                        paint.setStyle(Paint.Style.STROKE);
                        paint.setStrokeWidth(2);
                        paint.setColor(Color.BLACK);
                        canvas.drawRect((i * row_size) + row_index_size,
                                (j * col_size) + col_index_size,
                                (i * row_size) + row_size + row_index_size,
                                (j * col_size) + col_size + col_index_size,
                                paint);

                        // end, start 값을 지정하면, 그만큼의 길이가 안되는 셀은 출력할때 오류 발생
                        /* 셀의 값 그리기 */
                        paint.setColor(Color.BLACK);
                        paint.setStyle(Paint.Style.FILL);
                        if(excelload.length() > 7){
                            excelload_alpha = excelload.substring(0, 5) + "..";
                            canvas.drawText(excelload_alpha,
                                    0,
                                    7,
                                    (i * row_size) + (row_size / 2) + row_index_size,
                                    (j * col_size) + (col_size * 3 / 4) + col_index_size,
                                    paint);
                        }
                        else{
                            canvas.drawText(excelload,
                                    (i * row_size) + (row_size / 2) + row_index_size,
                                    (j * col_size) + (col_size * 3 / 4) + col_index_size,
                                    paint);
                        }
                    }
                    else{
                        /* 배경 셀 그리기 */
                        paint.setStyle(Paint.Style.FILL);
                        paint.setColor(Color.WHITE);
                        canvas.drawRect((i * row_size) + row_index_size,
                                (j * col_size) + col_index_size,
                                (i * row_size) + row_size + row_index_size,
                                (j * col_size) + col_size + col_index_size,
                                paint);

                        paint.setStyle(Paint.Style.STROKE);
                        paint.setStrokeWidth(1);
                        paint.setColor(Color.parseColor("#999999"));
                        canvas.drawRect((i * row_size) + row_index_size,
                                (j * col_size) + col_index_size,
                                (i * row_size) + row_size + row_index_size,
                                (j * col_size) + col_size + col_index_size,
                                paint);
                    }
                }
            }

            /* Paint Merge Cell */
            Range range[] = null;
            range = sheet.getMergedCells();

            for(Range rg : range){
                Log.d("merge topLeft", rg.getTopLeft().getContents());
                Log.d("merge col", Integer.toString(rg.getTopLeft().getColumn()));
                Log.d("merge row", Integer.toString(rg.getTopLeft().getRow()));
                Log.d("merge bottomRight", rg.getBottomRight().getContents());
                Log.d("merge row", Integer.toString(rg.getBottomRight().getColumn()));
                Log.d("merge col", Integer.toString(rg.getBottomRight().getRow()));

                Log.d("row_move_count", Integer.toString(row_move_count));
                Log.d("col_move_count", Integer.toString(col_move_count));

                /* drawRect 는 size 를 기준으로 하기 때문에, count 를 의미하는 값은 row, column 이 반대가 되어야 함
                * 4 * 15 의 canvas 를 그릴 땐, 움직이는 만큼 move_count 를 더해주어야 엑셀 시작값이 달라지지만
                * merge cell 의 경우 절대적인 위치이므로 이동한 만큼 다시 돌려놓아야 주므로 move count 를
                * 빼주어야(반대의 의미로 생각해야) 된다. */

                /* 병합 셀 배경 그리기 */
                paint.setStyle(Paint.Style.FILL);
                paint.setColor(Color.WHITE);
                canvas.drawRect(((rg.getTopLeft().getColumn() - col_move_count) * row_size) + row_index_size,
                        ((rg.getTopLeft().getRow() - row_move_count) * col_size) + col_index_size,
                        ((rg.getBottomRight().getColumn() - col_move_count) * row_size) + row_size + row_index_size,
                        ((rg.getBottomRight().getRow() - row_move_count) * col_size) + col_size + col_index_size,
                        paint);

                paint.setStyle(Paint.Style.STROKE);
                paint.setStrokeWidth(2);
                paint.setColor(Color.BLACK);
                canvas.drawRect(((rg.getTopLeft().getColumn() - col_move_count) * row_size) + row_index_size,
                        ((rg.getTopLeft().getRow() - row_move_count) * col_size) + col_index_size,
                        ((rg.getBottomRight().getColumn() - col_move_count) * row_size) + row_size + row_index_size,
                        ((rg.getBottomRight().getRow() - row_move_count) * col_size) + col_size + col_index_size,
                        paint);

                excelload = sheet.getCell(rg.getTopLeft().getColumn(), rg.getTopLeft().getRow()).getContents();

                int merge_TopLeft_row_size = (rg.getTopLeft().getColumn() - col_move_count) * row_size;
                int merge_TopLeft_col_size = (rg.getTopLeft().getRow() - row_move_count) * col_size;

                /* 병합된 topLeft, bottomRight Cell 은 병합된 쉘에 함께 포함되어 있기 때문에 +1 을 해주어야 함. */
                int merge_all_row_count = rg.getBottomRight().getColumn() - rg.getTopLeft().getColumn() + 1;
                int merge_all_col_count = rg.getBottomRight().getRow() - rg.getTopLeft().getRow() + 1;

                int merge_all_row_size = merge_all_row_count * row_size;
                int merge_all_col_size = merge_all_col_count * col_size;

                Log.d("merge_topLeft_col_size", Integer.toString(merge_TopLeft_col_size));
                Log.d("merge_all_col_size", Integer.toString(merge_all_col_size));
                Log.d("merge_text_loc_size", Integer.toString(merge_TopLeft_col_size + (merge_all_col_size / 2) + col_index_size));

                /* 병합된 데이터 영역중 topLeft 를 제외한 나머지 셀에다 같은 값 입력해주기 (write 작업을 하므로 원본을 따로 저장해둘지 생각해보기) */

                /* 병합 셀 데이터 그리기 */
                paint.setColor(Color.BLACK);
                paint.setStyle(Paint.Style.FILL);
                if(merge_all_col_count > 1){
                    if(excelload.length() > 7){
                        if(merge_all_row_count > 2 && excelload.length() > 13) {
                            excelload_alpha = excelload.substring(0, 10) + "...";
                            canvas.drawText(excelload_alpha,
                                    0,
                                    13,
                                    merge_TopLeft_row_size + (merge_all_row_size / 2) + row_index_size,
                                    merge_TopLeft_col_size + (merge_all_col_size / 2) + (col_size * 1 / 4) + col_index_size,
                                    paint);
                        }
                        else{
                            excelload_alpha = excelload.substring(0, 5) + "..";
                            canvas.drawText(excelload_alpha,
                                    0,
                                    7,
                                    merge_TopLeft_row_size + (merge_all_row_size / 2) + row_index_size,
                                    merge_TopLeft_col_size + (merge_all_col_size / 2) + (col_size * 1 / 4) + col_index_size,
                                    paint);
                        }
                    }
                    else{
                        canvas.drawText(excelload,
                                merge_TopLeft_row_size + (merge_all_row_size / 2) + row_index_size,
                                merge_TopLeft_col_size + (merge_all_col_size / 2) + (col_size * 1 / 4) + col_index_size,
                                paint);
                    }
                }
                else{
                    if(excelload.length() > 7) {
                        if(merge_all_row_count > 2 && excelload.length() > 13){
                            excelload_alpha = excelload.substring(0, 10) + "...";
                            canvas.drawText(excelload_alpha,
                                    0,
                                    13,
                                    merge_TopLeft_row_size + (merge_all_row_size / 2) + row_index_size,
                                    merge_TopLeft_col_size + (merge_all_col_size * 3 / 4) + col_index_size,
                                    paint);
                        }
                        else{
                            excelload_alpha = excelload.substring(0, 5) + "..";
                            canvas.drawText(excelload_alpha,
                                    0,
                                    7,
                                    merge_TopLeft_row_size + (merge_all_row_size / 2) + row_index_size,
                                    merge_TopLeft_col_size + (merge_all_col_size * 3 / 4) + col_index_size,
                                    paint);
                        }
                    }
                    else{
                        canvas.drawText(excelload,
                                merge_TopLeft_row_size + (merge_all_row_size / 2) + row_index_size,
                                merge_TopLeft_col_size + (merge_all_col_size * 3 / 4) + col_index_size,
                                paint);
                    }
                }
            }

            // index 를 나타내는 0행 0열 ----------------------------
            /* index Cell 배경 그리기 */
            paint.setStyle(Paint.Style.FILL);
            paint.setColor(Color.parseColor("#cccccc"));
            canvas.drawRect(0, 0, row_index_size, col_index_size, paint);

            for(k = 0; k < col_count; k++){
                canvas.drawRect((k * row_size) + row_index_size,
                        (0 * col_size),
                        (k * row_size) + row_size + row_index_size,
                        (0 * col_size) + col_index_size,
                        paint);
            }

            for(k = 0; k < row_count; k++){
                canvas.drawRect((0 * row_size),
                        (k * col_size) + col_index_size,
                        (0 * row_size) + row_index_size,
                        (k * col_size) + col_size + col_index_size,
                        paint);
            }

            paint.setStyle(Paint.Style.STROKE);
            paint.setColor(Color.parseColor("#999999"));
            for(k = 0; k < col_count; k++){
                canvas.drawRect((k * row_size) + row_index_size,
                        (0 * col_size),
                        (k * row_size) + row_size + row_index_size,
                        (0 * col_size) + col_index_size,
                        paint);
            }

            for(k = 0; k < row_count; k++){
                canvas.drawRect((0 * row_size),
                        (k * col_size) + col_index_size,
                        (0 * row_size) + row_index_size,
                        (k * col_size) + col_size + col_index_size,
                        paint);
            }


            /* index Cell 값 그리기 */
            paint.setColor(Color.BLACK);
            paint.setStyle(Paint.Style.FILL);
            for(k = 0; k < col_count; k++){
                canvas.drawText(Integer.toString(k + col_move_count),
                        (k * row_size) + (row_size / 2) + row_index_size,
                        (0 * col_size) + (col_size * 3 / 4),
                        paint);
            }

            for(k = 0; k < row_count; k++){
                canvas.drawText(Integer.toString(k + row_move_count),
                        (0 * row_size) + (row_index_size / 2),
                        (k * col_size) + (col_size * 3 / 4) + col_index_size,
                        paint);
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
