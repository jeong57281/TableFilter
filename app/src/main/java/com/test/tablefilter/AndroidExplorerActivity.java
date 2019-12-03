package com.test.tablefilter;

import android.app.AlertDialog;
import android.app.ListActivity;
import android.content.DialogInterface;
import android.content.Intent;
import android.os.Bundle;
import android.util.Log;
import android.view.View;
import android.widget.ArrayAdapter;
import android.widget.ListView;
import android.widget.TextView;

import androidx.annotation.Nullable;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

public class AndroidExplorerActivity extends ListActivity {
    /* 탐색기에 표시될 item 과 눌렀을 경우 이동할 경로 path 이다. */
    private List<String> item = null;
    private List<String> path = null;

    /* root 디렉토리 설정 */
    //private String root = Environment.getExternalStorageDirectory().getPath();
    private String root = "/sdcard/";

    /* 현재 경로를 담는 변수 myPath */
    private TextView myPath;

    @Override
    protected void onCreate(@Nullable Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.filemain);
        /* myPath 뷰어 설정 */
        myPath = (TextView) findViewById(R.id.path);
        getDir(root);
    }

    /* 주어진 argument(파일 경로)에 대한 탐색기 뷰 생성
    @param dirPath
     */

    private void getDir(String dirPath){
        /* myPath를 넘어온 매개변수로 설정 */
        myPath.setText("Location: " + dirPath);

        item = new ArrayList<String>();
        path = new ArrayList<String>();

        /* 주어진 주소값에 대한 File 객체 생성 및 하위
        디렉토리/파일 행렬 생성
         */

        File f = new File(dirPath);
        File[] files = f.listFiles();

        Log.d("dirPath", dirPath);
        Log.d("rootPath", root);
        /* item, path 추가 */
        if(!dirPath.equals(root)){
            /* root 폴더 아이템 및 경로 */
            item.add(root);
            path.add(root);

            /* 상위 디렉토리 및 경로 */
            item.add("../");
            path.add(f.getParent() + "/");
        }

        //for(int i = 0; i < files.length; i++){
        for(File file : files){
            //File file = files[i];
            path.add(file.getPath());

            if(file.isDirectory())
                item.add(file.getName() + "/");
            else
                item.add(file.getName());
        }

        /*
        ArrayAdapter를 통해 myPath에서의 파일 탐색기 뷰 생성
         */
        ArrayAdapter<String> fileList = new ArrayAdapter<String>(this, R.layout.row, item);
        setListAdapter(fileList);
    }

    @Override
    protected void onListItemClick(ListView l, View v, int position, long id) {
        /*
        현재 뷰가 root 디렉토리가 아닌 경우
        path.get(0) = root 경로
        path.get(1) = 상위 디렉토리 경로
        path.get(2) = getDir 에서 저장된 하위 디렉토리/파일 경로
         */
        File file = new File(path.get(position));

        /* 만약 디렉토리를 클락하였다면 */
        if(file.isDirectory()){
            /*
            file.canRead()는 파일에 대한 접근 권한을 확인하는
            메소드, canRead() 대신에 canExecute(), canRead()를
            사용할 수도 있다.
             */

            if(file.canRead()){
                getDir(path.get(position));
            }
            else{
                new AlertDialog.Builder(this)
                        .setIcon(R.drawable.button)
                        .setTitle("[" + file.getName() + "] folder can't be read!")
                        .setPositiveButton("OK",
                                new DialogInterface.OnClickListener() {
                    @Override
                    public void onClick(DialogInterface dialog, int which) {

                    }
                }).show();
            }
        }
        /* 만약 디렉토리가 아닌 파일을 클릭한 것이라면 */
        else{
            /* 해당 예제는 파일 클릭시 다이얼로그를 띄우는 action
            을 취한다. 의도하는 action 이 있다면 아래의 코드를 지우고
            원하는 코드로 대체한다.
             */

            String fileName;
            String extName;

            fileName = file.getName();

            extName = fileName.substring(fileName.length() - 3);
            if(extName.equals((String) "xls")){
                /* new Intent("현재 activity, 실행할 activity);
                * getApplicationContext() 와 AndroidExplorerActivity.this 는
                * 같은 의미일까요? */
                Intent xlsSetting = new Intent(AndroidExplorerActivity.this, xlsSetting.class);
                xlsSetting.putExtra("fileName", fileName);
                xlsSetting.putExtra("filePath", path.get(position));
                startActivity(xlsSetting);
            }

            /*
            new AlertDialog.Builder(this)
                    .setIcon(R.drawable.button)
                    .setTitle("[" + file.getName() + "]")
                    .setPositiveButton("OK",
                            new DialogInterface.OnClickListener() {
                                @Override
                                public void onClick(DialogInterface dialog, int which) {

                                }
                            }).show();
             */
        }
    }
}
