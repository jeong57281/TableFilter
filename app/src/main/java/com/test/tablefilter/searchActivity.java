package com.test.tablefilter;

import android.content.Intent;
import android.os.Bundle;
import android.util.Log;
import android.view.View;
import android.view.inputmethod.InputMethodManager;
import android.widget.ArrayAdapter;
import android.widget.AutoCompleteTextView;
import android.widget.Button;
import android.widget.ListView;
import android.widget.Spinner;
import android.widget.TextView;

import androidx.appcompat.app.AppCompatActivity;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Array;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

import javax.xml.transform.Result;

import jxl.Range;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;

public class searchActivity extends AppCompatActivity {
    // excel
    Workbook workbook = null;
    Sheet sheet = null;
    String[][] excelArray;
    Range range[] = null;
    userRange[] userRange = null;

    // view
    Button btnSearch;
    Button btnHome;
    TextView tvStatus; // tvTitle 은 키보드의 포커스를 맞추기 위해서 선언
    ListView result_ResultList;
    AutoCompleteTextView autoCompleteTextView;

    InputMethodManager imm;

    // etc
    boolean flag = false; // 일치하는 항목이 없을 경우, 안내 메세지를 위한 flag

    /* intent */
    Intent intent;

    // 전달받은 (표의)좌표값
    int row_start, col_start;
    int row_end, col_end;

    /* 실제 .xls sheet 의 최대 행, 열 개수 */
    int xls_row_max_count;
    int xls_col_max_count;

    // 전달받은 파일 경로
    String filePath;

    // ResultList
    List<String> SearchList;
    ArrayList<ResultListviewItem> ResultList;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.search_activity);

        btnSearch = (Button) findViewById(R.id.buttonSearch);
        tvStatus = (TextView) findViewById(R.id.tvStatus);
        result_ResultList = (ListView) findViewById(R.id.result_list);
        autoCompleteTextView = (AutoCompleteTextView) findViewById(R.id.autoCompleteTextView);

        // 출력 데이터를 저장하게 되는 리스트
        ResultList = new ArrayList<>();
        // 검색어 데이터를 저장하는 리스트
        SearchList = new ArrayList<>();

        // 리스트뷰와 리스트를 연결하기 위해 사용되는 어댑터
        /* ResultListviewAdapter ResultAdapter = new ResultListviewAdapter(this,
                R.layout.result_row, ResultList); search_Excel 메소드로 이동 */

        ArrayAdapter<String> SearchAdapter = new ArrayAdapter<>(this,
                android.R.layout.simple_list_item_1, SearchList);

        // excel 데이터 배열에 불러오기
        Excel();

        // 리스트뷰와 자동완성 텍스트 뷰의 어댑터를 지정해준다.
        // 자동완성 텍스트 뷰는 excel 파일에서 불러온 데이터를 1회성으로 리스트에 저장해 연결하지만,
        // 값을 검색했을 때 결과를 저장해주는 arraylist 는 매번 연결이 필요하다. 따라서 search_Excel 메소드로 이동
        /* result_ResultList.setAdapter(ResultAdapter); search_Excel 메소드로 이동 */
        autoCompleteTextView.setAdapter(SearchAdapter);

        // 검색어 중 하나를 랜덤으로 선택해 hint 로 추가
        Random random = new Random();
        autoCompleteTextView.setHint("예) " + SearchList.get(random.nextInt(SearchList.size())));

        // 확인 버튼을 누르면 키보드가 강제로 내려가도록 하기 위한 코드
        imm = (InputMethodManager) getSystemService(INPUT_METHOD_SERVICE);

        btnSearch.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                SearchExcel(autoCompleteTextView.getText().toString());
                imm.hideSoftInputFromWindow(autoCompleteTextView.getWindowToken(), 0);
            }
        });

        btnHome = (Button) findViewById(R.id.home_btn);
        btnHome.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                intent = new Intent(getApplicationContext(), MainActivity.class);
                startActivity(intent);
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
            WorkbookSettings ws = new WorkbookSettings();
            ws.setEncoding("Cp1252");
            workbook = Workbook.getWorkbook(file, ws);

            // 엑셀 파일의 첫 번째 시트 인식
            sheet = workbook.getSheet(0);

            // 엑셀 파일에 표만 존재하고 불순물이 없다는 전제하에 마지막(end)값은 받지 않아도 됨.
            // 값은 1부터 시작하므로 1을 빼줌.
            row_end = sheet.getRows();
            col_end = sheet.getColumns();

            // 불러온 엑셀 데이터의 최대 행, 열 값
            xls_row_max_count = sheet.getRows();
            xls_col_max_count = sheet.getColumns();
            Log.d("xls_row_max_count", Integer.toString(sheet.getRows()));
            Log.d("xls_col_max_count", Integer.toString(sheet.getColumns()));

            // xls 데이터를 담을 array 객체 생성
            excelArray = new String[xls_row_max_count][xls_col_max_count];

            for (int row = 0; row < xls_row_max_count; row++) {
                for (int col = 0; col < xls_col_max_count; col++) {
                    Log.d("table row", Integer.toString(row));
                    Log.d("table col", Integer.toString(col));
                    // getCell 은 열, 행 순 (좌표 개념)
                    String cell_value = sheet.getCell(col, row).getContents();
                    excelArray[row][col] = cell_value;
                }
            }

            // 검색 할 키워드(테이블의 0열)를 SearchList 에 저장
            for (int row = row_start + 1; row < row_end; row++) {
                String cell_value = excelArray[row][col_start];
                SearchList.add(cell_value);
            }

            // 병합 셀 Range 객체 range 에 저장
            range = sheet.getMergedCells();
            userRange = new userRange[range.length];

            int count = 0;
            for(Range rg : range){
                userRange[count] = new userRange(); // 객체 배열을 사용할 때 100 번 주의 !!!
                userRange[count].setTopLeftRow(rg.getTopLeft().getRow());
                userRange[count].setTopLeftColumn(rg.getTopLeft().getColumn());
                userRange[count].setBottomRightRow(rg.getBottomRight().getRow());
                userRange[count].setBottomRightColumn(rg.getBottomRight().getColumn());
                userRange[count].setTopLeftContents(rg.getTopLeft().getContents());
                count++;
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

            String rowItem = excelArray[row][col_start];

            // 대소문자 구분 없이 입력된 값을 받아 비교
            if (rowItem.equals(word) || rowItem.toLowerCase().equals(word)) {
                for(int col = col_start + 1; col < col_end; col++){
                    /* item 과 value 값을 찾는 부분에서, 만약 빈 값이 출력되면 해당 셀의 위치가 병합된 셀인지를
                    판단하고 여부에 따라 병합된 셀의 해당 값으로 반환
                    판단 기준 : row, col 의 쌍이 모든 병합 셀 영역과 비교하여 일치할 경우
                     */
                    String colItem = excelArray[row_start][col];
                    if(colItem.equals("")){
                        for(userRange userRange : userRange){
                            // mergeRow 와 mergeCol 은 count 개념, getRow 와 getColumn 은 셀 개념 : 반대
                            for(int mergeRow = userRange.getTopLeftRow(); mergeRow <= userRange.getBottomRightRow(); mergeRow++){
                                for(int mergeCol = userRange.getTopLeftColumn(); mergeCol <= userRange.getBottomRightColumn(); mergeCol++){
                                   /*
                                  Log.d("mergeRow", Integer.toString(mergeRow));
                                  Log.d("mergevs", Integer.toString(row));
                                  Log.d("mergeCol", Integer.toString(mergeCol));
                                  Log.d("mergevs", Integer.toString(col));
                                    */
                                    if(mergeRow == row_start && mergeCol == col){
                                        colItem = userRange.getTopLeftContents();
                                    }
                                }
                            }
                        }
                    }

                    String colValue = excelArray[row][col];
                    if(colValue.equals("")){
                       for(userRange userRange : userRange){
                           // mergeRow 와 mergeCol 은 count 개념, getRow 와 getColumn 은 셀 개념 : 반대
                           for(int mergeRow = userRange.getTopLeftRow(); mergeRow <= userRange.getBottomRightRow(); mergeRow++){
                               for(int mergeCol = userRange.getTopLeftColumn(); mergeCol <= userRange.getBottomRightColumn(); mergeCol++){
                                   /*
                                  Log.d("mergeRow", Integer.toString(mergeRow));
                                  Log.d("mergevs", Integer.toString(row));
                                  Log.d("mergeCol", Integer.toString(mergeCol));
                                  Log.d("mergevs", Integer.toString(col));
                                    */
                                  if(mergeRow == row && mergeCol == col){
                                      colValue = userRange.getTopLeftContents();
                                  }
                              }
                           }
                       }
                    }

                    Log.d("colItem", colItem);
                    Log.d("colValue", colValue);
                    ResultListviewItem resultItem = new ResultListviewItem(colItem, colValue);
                    Log.d("colResultItem", resultItem.getColItem());
                    Log.d("colResultValue", resultItem.getColValue());
                    ResultList.add(resultItem);
                }
                tvStatus.setText("항목을 찾았습니다.");
                flag = true;
            }
        }
        if(!flag){
            tvStatus.setText("일치하는 항목이 없습니다.");
        }

        ResultListviewAdapter ResultAdapter = new ResultListviewAdapter(this,
                R.layout.result_row, ResultList);
        result_ResultList.setAdapter(ResultAdapter);

        for(int i = 0; i < ResultList.size(); i++){
            Log.d("resultListItem", ResultList.get(i).getColItem());
            Log.d("resultListItem", ResultList.get(i).getColValue());
        }
    }
}
