package com.test.tablefilter;

import android.app.Activity;
import android.app.Dialog;
import android.content.Intent;
import android.content.SharedPreferences;
import android.graphics.Bitmap;
import android.graphics.Canvas;
import android.graphics.Color;
import android.graphics.Paint;
import android.graphics.Typeface;
import android.os.Bundle;
import android.os.Environment;
import android.view.View;
import android.widget.ArrayAdapter;
import android.widget.Button;
import android.widget.CheckBox;
import android.widget.ImageView;
import android.widget.Spinner;
import android.widget.TextView;
import android.widget.Toast;

import androidx.annotation.Nullable;
import androidx.appcompat.app.AppCompatActivity;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import jxl.Range;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class xlsSetting extends AppCompatActivity {

    // View
    Button btn;
    Button down_btn;
    Button up_btn;
    Button right_btn;
    Button left_btn;
    Button first_btn;
    Button last_btn;
    Button notice_btn;
    TextView tv_fileName;
    Spinner RowSpinner, ColSpinner;

    // spinner list
    List<String> RowList, ColList;

    // dialog
    Dialog dialog;
    TextView tvTitle, tvUsage, tvWarn;
    ImageView ivUsage, ivWarn;
    CheckBox dialogNoShowCb;

    String NoShowStatus;

    // intent
    Intent intent, intent2;
    String fileName;
    String filePath;

    // excel
    Sheet sheet = null;
    String excelload, excelload_alpha;
    String[][] excelArray;
    Range range[] = null;
    userRange[] userRange = null;

    // canvas
    Bitmap bitmap;
    Paint paint;
    ImageView imageView;
    Canvas canvas;

    // 반복문 변수
    int i, j, k;

    int text_size = 30;

    // xls file 미리보기를 구현할 캔버스 정보
    // 셀의 크기 지정
    int row_index_size = 50;
    int col_index_size = 40;
    int row_size = 160;
    int col_size = 40;

    // 셀의 개수 지정
    int row_count = 14;
    int col_count = 4;

    // 범위를 조정할 크기값
    int row_move_count = 0;
    int col_move_count = 0;

    // 실제 .xls sheet 의 최대 행, 열 개수
    int xls_row_max_count;
    int xls_col_max_count;

    // 미리보기에서 이동한 행, 열의 현재 위치
    int current_row_count;
    int current_col_count;

    // 입력받은 좌표값
    int row_start, col_start;

    @Override
    protected void onCreate(@Nullable Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.xls_setting);
        // ------------------------------------------------------------------------------
        // 도움말 dialog
        dialog = new Dialog(this, R.style.Dialog);
        dialog.setContentView(R.layout.custom_dialog);

        tvTitle = (TextView) dialog.findViewById(R.id.dialogTitle);
        tvTitle.setText("도움말");

        // 사용법 안내 image, textview
        ivUsage = (ImageView) dialog.findViewById(R.id.dialogUsageImage);
        ivUsage.setImageResource(R.drawable.spreadsheettable_hojadecalculo_9378);

        tvUsage = (TextView) dialog.findViewById(R.id.dialogUsageText);
        tvUsage.setText(R.string.Usage);

        // 잘못된 사용법 안내 image, textview
        tvWarn = (TextView) dialog.findViewById(R.id.dialogWarnText);
        tvWarn.setText(R.string.Warn);

        ivWarn = (ImageView) dialog.findViewById(R.id.dialogWarnImage);
        ivWarn.setImageResource(R.drawable.spreadsheettable_hojadecalculo_93);
        // ------------------------------------------------------------------------------
        dialogNoShowCb = (CheckBox) dialog.findViewById(R.id.cb_NoShow);

        SharedPreferences pref = getSharedPreferences("pref", Activity.MODE_PRIVATE);
        SharedPreferences.Editor editor = pref.edit();

        dialogNoShowCb.setChecked(Boolean.parseBoolean(pref.getString("NoShowStatus", "")));

        editor.putString("NoShowStatus", String.valueOf(dialogNoShowCb.isChecked()));
        editor.commit();

        NoShowStatus = pref.getString("NoShowStatus", "");

        if(NoShowStatus.equals((String) "false")){
           dialog.show();
        }

        Button dialogOkButton = (Button) dialog.findViewById(R.id.OkBtn);
        dialogOkButton.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                dialog.dismiss();

                if(dialogNoShowCb.isChecked()){
                    SharedPreferences pref = getSharedPreferences("pref", Activity.MODE_PRIVATE);
                    SharedPreferences.Editor editor = pref.edit();

                    editor.putString("NoShowStatus", String.valueOf(dialogNoShowCb.isChecked()));
                    editor.commit();

                    NoShowStatus = pref.getString("NoShowStatus", "");

                    Toast.makeText(getApplicationContext(), "도움말 버튼을 통해 다시 확인할 수 있습니다.", Toast.LENGTH_LONG).show();
                }
                else{
                    SharedPreferences pref = getSharedPreferences("pref", Activity.MODE_PRIVATE);
                    SharedPreferences.Editor editor = pref.edit();

                    editor.putString("NoShowStatus", String.valueOf(dialogNoShowCb.isChecked()));
                    editor.commit();

                    NoShowStatus = pref.getString("NoShowStatus", "");
                }
            }
        });
        // ------------------------------------------------------------------------------
        // 유의사항 출력 버튼
        notice_btn = (Button) findViewById(R.id.notice_btn);
        notice_btn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                dialog.show();
            }
        });
        // ------------------------------------------------------------------------------
        // intent
        intent = getIntent();
        filePath = intent.getExtras().getString("filePath");
        fileName = intent.getExtras().getString("fileName");

        tv_fileName = (TextView) findViewById(R.id.tv_fileName);
        tv_fileName.setText("파일명 : " + fileName);
        // ------------------------------------------------------------------------------
        // 외부 글꼴 지정
        Typeface tf = Typeface.createFromAsset(getAssets(), "AppleSDGothicNeoM.ttf");
        // ------------------------------------------------------------------------------
        // spinner
        RowSpinner = (Spinner) findViewById(R.id.rowSpinner);
        ColSpinner = (Spinner) findViewById(R.id.colSpinner);

        // spinner를 연결하기 위해 사용되는 어댑터
        RowList = new ArrayList<>();
        ColList = new ArrayList<>();

        ArrayAdapter<String> RowAdapter = new ArrayAdapter<>(this,
                android.R.layout.simple_list_item_1, RowList);

        ArrayAdapter<String> ColAdapter = new ArrayAdapter<>(this,
                android.R.layout.simple_list_item_1, ColList);
        // ------------------------------------------------------------------------------
        // excel 데이터 배열에 불러오기
        Excel();

        // Excel 함수에서 Row, Col List 값들이 저장됨.
        // 행, 열 의 spinner 에 List 를 연결
        RowSpinner.setAdapter(RowAdapter);
        ColSpinner.setAdapter(ColAdapter);

        // 미리보기 첫 호출
        Preview_Execl("none");
        // ------------------------------------------------------------------------------
        // 확인 버튼
        btn = (Button) findViewById(R.id.search_btn);
        btn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                intent2 = new Intent(getApplicationContext(), searchActivity.class);
                /* 파일 경로를 전달 */
                intent2.putExtra("filePath", filePath);

                // 값을 선택하지 않았을 경우 toast 경고 메세지 출력
                if(RowSpinner.getSelectedItem().toString().equals("선택") ||
                    ColSpinner.getSelectedItem().toString().equals("선택")){

                    Toast.makeText(getApplicationContext(), "값을 선택하세요.", Toast.LENGTH_LONG).show();
                }
                else{
                    /* 입력한 좌표값을 전달 */
                    row_start = Integer.parseInt(RowSpinner.getSelectedItem().toString()) - 1;
                    col_start = Integer.parseInt(ColSpinner.getSelectedItem().toString()) - 1;

                    intent2.putExtra("row_start", row_start);
                    intent2.putExtra("col_start", col_start);

                    /* 파일 경로명과, 좌표값을 recent.xls 에 기록 */
                    Recent_Database(fileName, filePath, row_start, col_start);

                    startActivity(intent2);
                }
            }
        });
        // ------------------------------------------------------------------------------
        // 엑셀 미리보기 이동버튼
        down_btn = (Button) findViewById(R.id.down_btn);
        down_btn.setTypeface(tf);
        down_btn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Preview_Execl("down");
            }
        });

        up_btn = (Button) findViewById(R.id.up_btn);
        up_btn.setTypeface(tf);
        up_btn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Preview_Execl("up");
            }
        });

        right_btn = (Button) findViewById(R.id.right_btn);
        right_btn.setTypeface(tf);
        right_btn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Preview_Execl("right");
            }
        });

        left_btn = (Button) findViewById(R.id.left_btn);
        left_btn.setTypeface(tf);
        left_btn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Preview_Execl("left");
            }
        });

        first_btn = (Button) findViewById(R.id.first_btn);
        first_btn.setTypeface(tf);
        first_btn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Preview_Execl("first");
            }
        });

        last_btn = (Button) findViewById(R.id.last_btn);
        last_btn.setTypeface(tf);
        last_btn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Preview_Execl("last");
            }
        });
    }

    public void Preview_Execl(String direction) {
        // bitmap 크기 설정, Canvas 에 연결
        bitmap = Bitmap.createBitmap(row_size * col_count + row_index_size,
                col_size * row_count + col_index_size,
                Bitmap.Config.ARGB_8888);
        canvas = new Canvas(bitmap);
        canvas.drawColor(Color.BLACK);
        // ------------------------------------------------------------------------------
        // 비트맵을 출력할 ImageView 와 연결
        imageView = (ImageView) findViewById(R.id.sample_image);
        imageView.setImageBitmap(bitmap);
        // ------------------------------------------------------------------------------
        // canvas 에 그리기 위한 도구인 paint 객체 선언
        paint = new Paint();
        paint.setColor(Color.RED);
        // ------------------------------------------------------------------------------
        // 사각형 스타일 설정 (테두리만 그리기)
        paint.setStyle(Paint.Style.FILL);
        paint.setStrokeWidth(1);
        // ------------------------------------------------------------------------------
        // Word Style
        paint.setColor(Color.BLACK);
        paint.setTextSize(text_size);
        paint.setTextAlign(Paint.Align.CENTER);
        // ------------------------------------------------------------------------------
        /* current_row_count, current_col_count는 현재의 최대 index 값을 보여주므로
        엑셀 시트의 최대 데이터 범위보다 작을 경우에만 증가시켜야 한다.
        같을 경우에 증가시키면 시트의 범위를 벗어나게 된다. */

        current_row_count = row_count + row_move_count;
        current_col_count = col_count + col_move_count;

        switch (direction) {
            case "down":
                if (current_row_count < xls_row_max_count)
                    row_move_count++;
                break;
            case "up":
                if (row_move_count > 0)
                    row_move_count--;
                break;
            case "right":
                if (current_col_count < xls_col_max_count)
                    col_move_count++;
                break;
            case "left":
                if (col_move_count > 0)
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

        // ------------------------------------------------------------------------------
        /* 4x15 사이즈의 엑셀 cell 이 미리보기에 출력된다. 4x15 보다 큰 엑셀 데이터만 사용하도록 권고. */

        for (i = 0; i < col_count; i++) {
            for (j = 0; j < row_count; j++) {
                excelload = excelArray[j + row_move_count][i + col_move_count];
                // 셀에 값이 있을 경우
                if (!excelload.equals("")) {
                    /* 배경 셀 그리기 */
                    paint.setStyle(Paint.Style.FILL);
                    paint.setColor(Color.rgb(252, 252, 255));
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
                    if (excelload.length() > 7) {
                        excelload_alpha = excelload.substring(0, 5) + "..";
                        canvas.drawText(excelload_alpha,
                                0,
                                7,
                                (i * row_size) + (row_size / 2) + row_index_size,
                                (j * col_size) + (col_size * 3 / 4) + col_index_size,
                                paint);
                    } else {
                        canvas.drawText(excelload,
                                (i * row_size) + (row_size / 2) + row_index_size,
                                (j * col_size) + (col_size * 3 / 4) + col_index_size,
                                paint);
                    }
                } else {
                    /* 배경 셀 그리기 */
                    paint.setStyle(Paint.Style.FILL);
                    paint.setColor(Color.rgb(252, 252, 255));
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

        // ------------------------------------------------------------------------------
        /* Paint Merge Cell */
        for (userRange rg : userRange) {
            /* drawRect 는 size 를 기준으로 하기 때문에, count 를 의미하는 값은 row, column 이 반대가 되어야 함
             * 4 * 15 의 canvas 를 그릴 땐, 움직이는 만큼 move_count 를 더해주어야 엑셀 시작값이 달라지지만
             * merge cell 의 경우 절대적인 위치이므로 이동한 만큼 다시 돌려놓아야 주므로 move count 를
             * 빼주어야(반대의 의미로 생각해야) 된다. */

            /* 병합 셀 배경 그리기 */
            paint.setStyle(Paint.Style.FILL);
            paint.setColor(Color.rgb(252, 252, 255));
            canvas.drawRect(((rg.getTopLeftColumn() - col_move_count) * row_size) + row_index_size,
                    ((rg.getTopLeftRow() - row_move_count) * col_size) + col_index_size,
                    ((rg.getBottomRightColumn() - col_move_count) * row_size) + row_size + row_index_size,
                    ((rg.getBottomRightRow() - row_move_count) * col_size) + col_size + col_index_size,
                    paint);

            paint.setStyle(Paint.Style.STROKE);
            paint.setStrokeWidth(2);
            paint.setColor(Color.BLACK);
            canvas.drawRect(((rg.getTopLeftColumn() - col_move_count) * row_size) + row_index_size,
                    ((rg.getTopLeftRow() - row_move_count) * col_size) + col_index_size,
                    ((rg.getBottomRightColumn() - col_move_count) * row_size) + row_size + row_index_size,
                    ((rg.getBottomRightRow() - row_move_count) * col_size) + col_size + col_index_size,
                    paint);


            excelload = excelArray[rg.getTopLeftRow()][rg.getTopLeftColumn()];

            int merge_TopLeft_row_size = (rg.getTopLeftColumn() - col_move_count) * row_size;
            int merge_TopLeft_col_size = (rg.getTopLeftRow() - row_move_count) * col_size;

            int merge_all_row_count = rg.getMergeRowCount();
            int merge_all_col_count = rg.getMergeColCount();

            int merge_all_row_size = merge_all_row_count * row_size;
            int merge_all_col_size = merge_all_col_count * col_size;

            /* 병합된 데이터 영역중 topLeft 를 제외한 나머지 셀에다 같은 값 입력해주기 (write 작업을 하므로 원본을 따로 저장해둘지 생각해보기) */

            /* 병합 셀 데이터 그리기 */
            paint.setColor(Color.BLACK);
            paint.setStyle(Paint.Style.FILL);
            if (merge_all_col_count > 1) {
                if (excelload.length() > 7) {
                    if (merge_all_row_count > 2 && excelload.length() > 13) {
                        excelload_alpha = excelload.substring(0, 10) + "...";
                        canvas.drawText(excelload_alpha,
                                0,
                                13,
                                merge_TopLeft_row_size + (merge_all_row_size / 2) + row_index_size,
                                merge_TopLeft_col_size + (merge_all_col_size / 2) + (col_size * 1 / 4) + col_index_size,
                                paint);
                    } else {
                        excelload_alpha = excelload.substring(0, 5) + "..";
                        canvas.drawText(excelload_alpha,
                                0,
                                7,
                                merge_TopLeft_row_size + (merge_all_row_size / 2) + row_index_size,
                                merge_TopLeft_col_size + (merge_all_col_size / 2) + (col_size * 1 / 4) + col_index_size,
                                paint);
                    }
                } else {
                    canvas.drawText(excelload,
                            merge_TopLeft_row_size + (merge_all_row_size / 2) + row_index_size,
                            merge_TopLeft_col_size + (merge_all_col_size / 2) + (col_size * 1 / 4) + col_index_size,
                            paint);
                }
            } else {
                if (excelload.length() > 7) {
                    if (merge_all_row_count > 2 && excelload.length() > 13) {
                        excelload_alpha = excelload.substring(0, 10) + "...";
                        canvas.drawText(excelload_alpha,
                                0,
                                13,
                                merge_TopLeft_row_size + (merge_all_row_size / 2) + row_index_size,
                                merge_TopLeft_col_size + (merge_all_col_size * 3 / 4) + col_index_size,
                                paint);
                    } else {
                        excelload_alpha = excelload.substring(0, 5) + "..";
                        canvas.drawText(excelload_alpha,
                                0,
                                7,
                                merge_TopLeft_row_size + (merge_all_row_size / 2) + row_index_size,
                                merge_TopLeft_col_size + (merge_all_col_size * 3 / 4) + col_index_size,
                                paint);
                    }
                } else {
                    canvas.drawText(excelload,
                            merge_TopLeft_row_size + (merge_all_row_size / 2) + row_index_size,
                            merge_TopLeft_col_size + (merge_all_col_size * 3 / 4) + col_index_size,
                            paint);
                }
            }
        }

        // ------------------------------------------------------------------------------
        // index 를 나타내는 0행 0열
        /* index Cell 배경 그리기 */
        paint.setStyle(Paint.Style.FILL);
        paint.setColor(Color.rgb(237, 237, 255));
        canvas.drawRect(0, 0, row_index_size, col_index_size, paint);

        for (k = 0; k < col_count; k++) {
            canvas.drawRect((k * row_size) + row_index_size,
                    (0 * col_size),
                    (k * row_size) + row_size + row_index_size,
                    (0 * col_size) + col_index_size,
                    paint);
        }

        for (k = 0; k < row_count; k++) {
            canvas.drawRect((0 * row_size),
                    (k * col_size) + col_index_size,
                    (0 * row_size) + row_index_size,
                    (k * col_size) + col_size + col_index_size,
                    paint);
        }

        paint.setStyle(Paint.Style.STROKE);
        paint.setColor(Color.parseColor("#999999"));
        for (k = 0; k < col_count; k++) {
            canvas.drawRect((k * row_size) + row_index_size,
                    (0 * col_size),
                    (k * row_size) + row_size + row_index_size,
                    (0 * col_size) + col_index_size,
                    paint);
        }

        for (k = 0; k < row_count; k++) {
            canvas.drawRect((0 * row_size),
                    (k * col_size) + col_index_size,
                    (0 * row_size) + row_index_size,
                    (k * col_size) + col_size + col_index_size,
                    paint);
        }

        canvas.drawRect(0, 0, row_index_size, col_index_size, paint);

        /* index Cell 값 그리기 */
        paint.setColor(Color.BLACK);
        paint.setStyle(Paint.Style.FILL);
        for (k = 0; k < col_count; k++) {
            canvas.drawText(Integer.toString(k + col_move_count + 1),
                    (k * row_size) + (row_size / 2) + row_index_size,
                    (0 * col_size) + (col_size * 3 / 4),
                    paint);
        }

        for (k = 0; k < row_count; k++) {
            canvas.drawText(Integer.toString(k + row_move_count + 1),
                    (0 * row_size) + (row_index_size / 2),
                    (k * col_size) + (col_size * 3 / 4) + col_index_size,
                    paint);
        }
    }

    Workbook writableWorkbook;
    public void Excel() {
        try {
            File file = new File(filePath);

            /* jxl encoding setting : utf-8 */
            WorkbookSettings ws = new WorkbookSettings();
            ws.setEncoding("Cp1252");
            // 이미 writableWorkbook 을 사용하고 있었네..
            writableWorkbook = Workbook.getWorkbook(file, ws);

            // 엑셀 파일의 첫 번째 시트 인식
            sheet = writableWorkbook.getSheet(0);

            // 불러온 엑셀 데이터의 최대 행, 열 값
            xls_row_max_count = sheet.getRows();
            xls_col_max_count = sheet.getColumns();

            // xls 데이터를 담을 array 객체 생성
            excelArray = new String[xls_row_max_count][xls_col_max_count];

            for (int row = 0; row < xls_row_max_count; row++) {
                for (int col = 0; col < xls_col_max_count; col++) {
                    // getCell 은 열, 행 순 (좌표 개념)
                    String cell_value = sheet.getCell(col, row).getContents();
                    excelArray[row][col] = cell_value;
                }
            }

            RowList.add("선택");
            ColList.add("선택");

            for (int row = 0; row < xls_row_max_count; row++) {
                RowList.add(Integer.toString(row + 1));
            }

            for (int col = 0; col < xls_col_max_count; col++) {
                ColList.add(Integer.toString(col + 1));
            }

            // 병합 셀 Range 객체 range 에 저장
            range = sheet.getMergedCells();
            userRange = new userRange[range.length];

            int count = 0;
            for (Range rg : range) {
                userRange[count] = new userRange(); // 객체 배열을 사용할 때 100 번 주의 !!!
                userRange[count].setTopLeftRow(rg.getTopLeft().getRow());
                userRange[count].setTopLeftColumn(rg.getTopLeft().getColumn());
                userRange[count].setBottomRightRow(rg.getBottomRight().getRow());
                userRange[count].setBottomRightColumn(rg.getBottomRight().getColumn());
                userRange[count].setTopLeftContents(rg.getTopLeft().getContents());
                count++;
            }

        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        } finally {
            writableWorkbook.close();
        }
    }

    Workbook wb;
    public void Recent_Database(String fileName, String filePath, int row_start, int col_start) {
        WritableWorkbook writableWorkbook = null;
        WritableSheet excelSheet = null;
        String saveFile = "Recent.xls";
        String saveFolder = "/Table Filter";
        String sheetName = "Sheet";
        try {
            File saveFolderPath = new File(Environment.getExternalStorageDirectory().getAbsolutePath() + saveFolder);

            if (!saveFolderPath.exists()) {
                saveFolderPath.mkdir();
            }

            File file = new File(saveFolderPath, saveFile);

            /* 일반 워크북을 하나 사용해서 먼저 읽은 다음, WritableWorkBook에 경로명과 함께 작성해주면
             기존의 파일에 이어서 작성이 가능하다고 하다. 무슨 차이길래? */
            try{
                wb = Workbook.getWorkbook(file); // wb is WorkBook
                writableWorkbook = Workbook.createWorkbook(file, wb);
            } catch (FileNotFoundException e){
                writableWorkbook = Workbook.createWorkbook(file);
            }

            // 만약 시트가 하나도 없다면 0번째 시트를 생성, 그렇지 않다면 해당 sheet 를 선택한다.
            if(writableWorkbook.getNumberOfSheets() == 0) {
                excelSheet = writableWorkbook.createSheet(sheetName, 0);
            }
            else {
                excelSheet = writableWorkbook.getSheet(sheetName);
            }

            String data[] = new String[]
                    {fileName, filePath, Integer.toString(row_start), Integer.toString(col_start)};

            int xls_row_max_count = excelSheet.getRows();

            // 만약 같은 파일이름의 항목이 있다면 해당 행에 덮어쓰기한다.
            int existRow = 0;
            boolean existFlag = false;
            for (int row = 0; row < xls_row_max_count; row++){
                if(fileName.equals(excelSheet.getCell(0, row).getContents())){
                    existRow = excelSheet.getCell(0, row).getRow();
                    existFlag = true;
                }
            }

            if(existFlag == true){
                for (int col = 0; col < 4; col++) {
                    Label label = new Label(col, existRow, data[col]);
                    excelSheet.addCell(label);
                }
            }
            else{
                for (int col = 0; col < 4; col++) {
                    Label label = new Label(col, xls_row_max_count, data[col]);
                    excelSheet.addCell(label);
                }
            }

            if(wb != null){
                wb.close();
            }

            writableWorkbook.write();
            writableWorkbook.close();

        } catch (IOException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        }
    }
}