package com.test.tablefilter;

import android.Manifest;
import android.app.Activity;
import android.app.Dialog;
import android.content.DialogInterface;
import android.content.Intent;
import android.content.SharedPreferences;
import android.content.pm.PackageInfo;
import android.content.pm.PackageManager;
import android.content.res.AssetManager;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.provider.Settings;
import android.util.Log;
import android.view.View;
import android.widget.AdapterView;
import android.widget.Button;
import android.widget.ListView;
import android.widget.TextView;
import android.widget.Toast;

import androidx.appcompat.app.AlertDialog;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.content.ContextCompat;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class MainActivity extends AppCompatActivity implements RecentListviewAdapter.ListBtnClickListener{

    // 최근목록 listview
    ListView recent_list;
    ArrayList<RecentListviewItem> RecentList;

    // 최근목록 클릭 시 전달 param
    String filePath;
    int row_start;
    int col_start;

    // dialog
    Dialog dialog, sample_dialog;

    // 최근 목록 삭제 시 전달되는 listview position 값
    int delete_position;

    // 첫 실행되는 sample file 을 위한 변수
    String version;
    String check_version, check_status;

    Workbook wb;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        RecentList = new ArrayList<>();

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
        // 첫 실행시에만 sample file 을 마련 - version 차이를 이용
        try {
            PackageInfo i = getPackageManager().getPackageInfo(getPackageName(), 0);
            version = i.versionName;
        } catch (PackageManager.NameNotFoundException e) {
            version = "";
        }

        SharedPreferences pref = getSharedPreferences("pref", Activity.MODE_PRIVATE);
        SharedPreferences.Editor editor = pref.edit();
        editor.putString("check_version", version);

        editor.commit();

        check_version = pref.getString("check_version", "");
        check_status = pref.getString("check_status", "");
        // ------------------------------------------------------------------------------
        // sample 을 추가할 것인지 묻는 dialog

        sample_dialog = new Dialog(this, R.style.Dialog);
        sample_dialog.setContentView(R.layout.sample_dialog);

        if(!check_version.equals(check_status)){
            TextView tv_sampleTitle = (TextView) sample_dialog.findViewById(R.id.tv_sampleTitle);
            tv_sampleTitle.setText("샘플 추가");

            TextView tv_sampleMainText = (TextView) sample_dialog.findViewById(R.id.tv_sampleMainText);
            tv_sampleMainText.setText("샘플 파일을 추가하시겠습니까?");

            sample_dialog.show();
        }

        Button sample_OKBtn = (Button) sample_dialog.findViewById(R.id.sample_OkBtn);
        sample_OKBtn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                int permissionStorageRead = ContextCompat.checkSelfPermission(getApplicationContext(), Manifest.permission.READ_EXTERNAL_STORAGE);
                int permissionStorageWrit = ContextCompat.checkSelfPermission(getApplicationContext(), Manifest.permission.WRITE_EXTERNAL_STORAGE);
                if(permissionStorageRead == PackageManager.PERMISSION_DENIED && permissionStorageWrit == PackageManager.PERMISSION_DENIED){
                    CALLDialog();
                }
                else{
                    Create_Sample_File();
                    Create_Recent_List();
                }
                sample_dialog.cancel();
            }
        });

        Button sample_CancelBtn = (Button) sample_dialog.findViewById(R.id.sample_CancelBtn);
        sample_CancelBtn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                sample_dialog.cancel();
            }
        });
        // ------------------------------------------------------------------------------
        // 최근목록 삭제 dialog
        dialog = new Dialog(this, R.style.Dialog);
        dialog.setContentView(R.layout.confirm_dialog);

        Button confirm_OKBtn = (Button) dialog.findViewById(R.id.confirm_OkBtn);
        confirm_OKBtn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Delete_Recent_List(delete_position);
                Create_Recent_List();
                dialog.cancel();
            }
        });

        Button confirm_CancelBtn = (Button) dialog.findViewById(R.id.confirm_CancelBtn);
        confirm_CancelBtn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                dialog.cancel();
            }
        });
        // ------------------------------------------------------------------------------
        // 최근 목록 클릭 시 동작
        recent_list = (ListView) findViewById(R.id.recent_list);
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
        // xls 불러오기 시 권한 체크
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
        // 메인 페이지(activity_main)가 새로 호출될 경우 최근목록 업데이트
        Create_Recent_List();
    }

    // sample file 어플 첫 실행시에만 sample file 마련
    public void Create_Sample_File(){
        // ------------------------------------------------------------------------------
        // 공통
        String saveFolder = "/Table Filter";
        String exampleSaveFolder = "/example";

        // Recent.xls 기록을 위한 변수
        WritableWorkbook writableWorkbook = null;
        WritableSheet excelSheet = null;
        String sheetName = "Sheet";
        String sheetFile = "Recent.xls";

        // asset 파일 복사를 위한 변수
        AssetManager assetManager = getAssets();
        String[] assets;
        InputStream is = null;
        FileOutputStream fos = null;
        byte[] buf = new byte[1024];

        try{
            // Table Filter, Table Filter/example 폴더 생성
            File saveFolderPath = new File(Environment.getExternalStorageDirectory().getAbsolutePath()
                    + saveFolder);

            if (!saveFolderPath.exists()) {
                saveFolderPath.mkdir();
            }

            File exampleSaveFolderPath = new File(Environment.getExternalStorageDirectory().getAbsolutePath()
                    + saveFolder
                    + exampleSaveFolder);

            if (!exampleSaveFolderPath.exists()) {
                exampleSaveFolderPath.mkdir();
            }

            File excelfile = new File(saveFolderPath, sheetFile);

            // asset 폴더의 모든 파일을 list 로 저장
            assets = assetManager.list("example");

            for(String element : assets){
                // 복사할 asset 파일 설정
                File copyfile = new File(Environment.getExternalStorageDirectory().getAbsolutePath()
                        + saveFolder
                        + exampleSaveFolder, element);

                // open 함수를 사용할 때, 맨 앞 / 는 생략해야 한다.
                is = assetManager.open("example/" + element);
                fos = new FileOutputStream(copyfile);

                while(is.read(buf) > 0){
                    fos.write(buf);
                }

                fos.close();
                is.close();
            }

            // Recent.xls 기록
            try{
                wb = Workbook.getWorkbook(excelfile); // wb is WorkBook
                writableWorkbook = Workbook.createWorkbook(excelfile, wb);
            } catch (FileNotFoundException e){
                writableWorkbook = Workbook.createWorkbook(excelfile);
            }

            if(writableWorkbook.getNumberOfSheets() == 0) {
                excelSheet = writableWorkbook.createSheet(sheetName, 0);
            }
            else {
                excelSheet = writableWorkbook.getSheet(sheetName);
            }

            int count = 0;
            int[][] index = {{4, 2}, {2, 2}, {4, 1}, {3, 1}};
            for(String element : assets) {
                String data[] = new String[]
                        {element, Environment.getExternalStorageDirectory().getAbsolutePath()
                                + saveFolder
                                + exampleSaveFolder
                                + "/" + element, Integer.toString(index[count][0]), Integer.toString(index[count][1])};

                int xls_row_max_count = excelSheet.getRows();

                // 만약 같은 파일이름의 항목이 있다면 해당 행에 덮어쓰기한다.
                int existRow = 0;
                boolean existFlag = false;
                for (int row = 0; row < xls_row_max_count; row++){
                    if(element.equals(excelSheet.getCell(0, row).getContents())){
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

                count++;
            }

            if(wb != null){
                wb.close();
            }

            writableWorkbook.write();
            writableWorkbook.close();

        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e){
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        }
        // ------------------------------------------------------------------------------
        // 실행 시, 현재 어플리케이션의 버전을 저장
        try {
            PackageInfo i = getPackageManager().getPackageInfo(getPackageName(), 0);
            version = i.versionName;
        } catch (PackageManager.NameNotFoundException e) {
            version = "";
        }

        SharedPreferences pref = getSharedPreferences("pref", Activity.MODE_PRIVATE);
        SharedPreferences.Editor editor = pref.edit();
        editor.putString("check_status", version);
        editor.commit();
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

            // 최근목록 삭제 부분
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

    // 사용자에게 권한요청 요구를 위한 다이어로그를 생성
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

    public void Confirm_Delete_List(int position){
        TextView tv_confirmTitle = (TextView) dialog.findViewById(R.id.tv_confirmTitle);
        tv_confirmTitle.setText("최근목록 삭제");

        TextView tv_confirmMainText = (TextView) dialog.findViewById(R.id.tv_confirmMainText);
        tv_confirmMainText.setText("최근목록에서 '" + RecentList.get(position).getRecentFileName() + "'을 삭제하시겠습니까?");

        dialog.show();
    }

    // Adapter 에서 구현한 row item 의 버튼 리스너
    @Override
    public void onListBtnClick(int position) {
        delete_position = position;
        Confirm_Delete_List(position);
    }
}
