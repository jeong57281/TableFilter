package com.test.tablefilter;

import android.Manifest;
import android.content.DialogInterface;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.provider.Settings;
import android.view.View;
import android.widget.AdapterView;
import android.widget.Button;
import android.widget.ListView;
import android.widget.Toast;

import androidx.appcompat.app.AlertDialog;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.content.ContextCompat;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class MainActivity extends AppCompatActivity implements RecentListviewAdapter.ListBtnClickListener {

    ListView recent_list;

    ArrayList<RecentListviewItem> RecentList;

    String filePath;
    int row_start;
    int col_start;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        // ------------------------------------------------------------------------------
        /* Storage READ 권한 획득
        api 23 이하에서는 항상 권한이 부여되므로 필요가 없다. */
        if(Build.VERSION.SDK_INT > 23){
            requestPermissions(new String[]
                            {Manifest.permission.READ_EXTERNAL_STORAGE,
                                    Manifest.permission.WRITE_EXTERNAL_STORAGE},
                    2);
        }
        // ------------------------------------------------------------------------------
        // 출력 데이터를 저장하게 되는 리스트
        RecentList = new ArrayList<>();

        recent_list = (ListView) findViewById(R.id.recent_list);

        Create_Recent_List();

        recent_list.setOnItemClickListener(new AdapterView.OnItemClickListener() {
            @Override
            public void onItemClick(AdapterView<?> parent, View view, int position, long id) {
                // 최근 열어본 xls 파일의 경로와 정보가 함께 저장되어 리스트로 출력해야 함
                if(Read_Recent_File(position)){
                    Intent intent = new Intent(getApplicationContext(), searchActivity.class);

                    intent.putExtra("filePath", filePath);
                    intent.putExtra("row_start", row_start);
                    intent.putExtra("col_start", col_start);

                    startActivity(intent);
                }
                else{
                    Delete_Recent_List(position);
                    Create_Recent_List();
                    Toast.makeText(getApplicationContext(), "파일이 삭제되었거나 변조되었습니다.", Toast.LENGTH_LONG).show();
                }
            }
        });
        // ------------------------------------------------------------------------------
        Button file_btn = (Button) findViewById(R.id.file_button);
        file_btn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                int permissionStorageRead = ContextCompat.checkSelfPermission(getApplicationContext(), Manifest.permission.READ_EXTERNAL_STORAGE);
                int permissionStorageWrit = ContextCompat.checkSelfPermission(getApplicationContext(), Manifest.permission.WRITE_EXTERNAL_STORAGE);
                if(permissionStorageRead == PackageManager.PERMISSION_DENIED && permissionStorageWrit == PackageManager.PERMISSION_DENIED){
                    CALLDialog();
                }
                else{
                    Intent file_intent_act = new Intent(getApplicationContext(), AndroidExplorerActivity.class);
                    startActivity(file_intent_act);
                }
            }
        });
        // ------------------------------------------------------------------------------
    }

    //사용자에게 권한요청 요구를 위한 다이어로그를 생성
    public void CALLDialog() {
        AlertDialog.Builder alertDialog = new AlertDialog.Builder(this);

        alertDialog.setTitle("앱 권한");
        alertDialog.setMessage("해당 앱의 원할한 기능을 이용하시려면 애플리케이션 정보>권한> 에서 모든 권한을 허용해 주십시오");

        alertDialog.setPositiveButton("권한설정",
                new DialogInterface.OnClickListener() {
                    public void onClick(DialogInterface dialog, int which) {
                        Intent intent = new Intent(Settings.ACTION_APPLICATION_DETAILS_SETTINGS).
                                setData(Uri.parse("package:" + getApplicationContext().getPackageName()));
                        startActivity(intent);
                        dialog.cancel();
                    }
                });
        alertDialog.setNegativeButton("취소",
                new DialogInterface.OnClickListener() {
                    public void onClick(DialogInterface dialog, int which) {
                        dialog.cancel();
                    }
                });

        alertDialog.show();
    }

    // Adapter 에서 구현한 row item 의 버튼 리스너
    @Override
    public void onListBtnClick(int position) {
        Delete_Recent_List(position);
        Create_Recent_List();
    }

    public boolean Read_Recent_File(int position){
        Sheet sheet = null;
        String saveFile = "Recent.xls";
        String saveFolder = "/Table Filter";
        try {
            File saveFolderPath = new File(Environment.getExternalStorageDirectory().getAbsolutePath() + saveFolder);

            File file = new File(saveFolderPath, saveFile);

            wb = Workbook.getWorkbook(file); // wb is WorkBook

            sheet = wb.getSheet(0);

            filePath = sheet.getCell(1, position).getContents();
            row_start = Integer.parseInt(sheet.getCell(2, position).getContents());
            col_start = Integer.parseInt(sheet.getCell(3, position).getContents());

            // 해당 경로에 파일이 존재하지 않는다면
            File f = new File(filePath);
            if(!f.isFile()){
                return false;
            }

            wb.close();

        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        }

        return true; // 정상적일 경우 true 반환
    }

    Workbook wb;
    public void Delete_Recent_List(int position){
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

            excelSheet.removeRow(position);

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

    Workbook workbook;
    public void Create_Recent_List(){
        RecentList.clear();
        Sheet sheet = null;
        String saveFile = "Recent.xls";
        String saveFolder = "/Table Filter";
        try {
            File saveFolderPath = new File(Environment.getExternalStorageDirectory().getAbsolutePath() + saveFolder);

            if (!saveFolderPath.exists()) {
                saveFolderPath.mkdir();
            }

            File file = new File(saveFolderPath, saveFile);

            workbook = Workbook.getWorkbook(file);

            sheet = workbook.getSheet(0);

            int xls_row_max_count = sheet.getRows();

            for(int row = 0; row < xls_row_max_count; row++){
                String fileItem = sheet.getCell(0, row).getContents();
                RecentListviewItem recentItem = new RecentListviewItem(fileItem);
                RecentList.add(recentItem);
            }

            RecentListviewAdapter RecentAdapter = new RecentListviewAdapter(this,
                    R.layout.recent_row, RecentList, this);
            recent_list.setAdapter(RecentAdapter);

            workbook.close();

        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        }
    }
}
