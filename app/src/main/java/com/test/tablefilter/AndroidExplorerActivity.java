package com.test.tablefilter;

import android.app.ListActivity;
import android.content.Intent;
import android.os.Bundle;
import android.view.View;
import android.widget.ListView;
import android.widget.TextView;
import android.widget.Toast;

import androidx.annotation.Nullable;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

public class AndroidExplorerActivity extends ListActivity {
    // 탐색기에 표시될 item 과 눌렀을 경우 이동할 경로 path 이다.
    private ArrayList<ListviewItem> item = null;
    private List<String> path = null;

    // root 디렉토리 설정
    //private String root = Environment.getExternalStorageDirectory().getPath();
    private String root = "/sdcard/";

    // 현재 경로를 담는 변수 myPath
    private TextView myPath;

    @Override
    protected void onCreate(@Nullable Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.filemain);
        /* myPath 뷰어 설정 */
        //myPath = (TextView) findViewById(R.id.path);
        getDir(root);
    }

    /* 주어진 argument(파일 경로)에 대한 탐색기 뷰 생성
    @param dirPath
     */
    private void getDir(String dirPath){
        // myPath를 넘어온 매개변수로 설정
        //myPath.setText("Location: " + dirPath);

        item = new ArrayList<ListviewItem>();
        path = new ArrayList<String>();

        // 주어진 주소값에 대한 File 객체 생성 및 하위 디렉토리,파일 리스트 생성

        File f = new File(dirPath);
        File[] files = f.listFiles();

        // item, path 추가
        if(!dirPath.equals(root)){
            // /sdcard/, ../ 대신 문자 사용
            // root 폴더 아이템 및 경로
            ListviewItem rootItem = new ListviewItem(R.drawable.folder_open_icon, "홈으로");
            item.add(rootItem);
            path.add(root);

            // 상위 디렉토리 및 경로
            ListviewItem backItem= new ListviewItem(R.drawable.folder_open_icon, "이전 폴더");
            item.add(backItem);
            path.add(f.getParent() + "/");
        }

        ListviewItem fileItem[] = new ListviewItem[files.length];
        int count = 0;
        for(File file : files){
            path.add(file.getPath());

            String fileName;

            fileName = file.getName();

            if(file.isDirectory()) {
                fileItem[count] = new ListviewItem(R.drawable.folder_open_icon, fileName);
                item.add(fileItem[count]);
            }
            else {
                if(fileName.substring(fileName.length() - 3).equals("xls")){
                    fileItem[count] = new ListviewItem(R.drawable.file_excel, fileName);
                }
                else{
                    fileItem[count] = new ListviewItem(R.drawable.file, fileName);
                }

                item.add(fileItem[count]);
            }
            count++;
        }

        ListviewAdapter fileList = new ListviewAdapter(this, R.layout.row, item);
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
            /* file.canRead()는 파일에 대한 접근 권한을 확인하는
            메소드, canRead() 대신에 canExecute(), canRead()를 사용할 수도 있다. */
            if(file.canRead()){
                getDir(path.get(position));
            }
            else{
                Toast.makeText(getApplicationContext(), "해당 폴더를 확인할 수 없습니다.", Toast.LENGTH_SHORT).show();
            }
        }
        /* 만약 디렉토리가 아닌 파일을 클릭한 것이라면 */
        else{
            String fileName;
            String ext3Name, ext4Name;

            fileName = file.getName();

            ext3Name = fileName.substring(fileName.length() - 3);
            ext4Name = fileName.substring(fileName.length() - 4);
            if(fileName.equals("Recent.xls")){
                Toast.makeText(getApplicationContext(), "해당 파일은 열어볼 수 없습니다.", Toast.LENGTH_SHORT).show();
            }
            else if(ext3Name.equals((String) "xls")){
                /* new Intent("현재 activity, 실행할 activity);
                * getApplicationContext() 와 AndroidExplorerActivity.this 는
                * 같은 의미일까요? */
                Intent xlsSetting = new Intent(AndroidExplorerActivity.this, xlsSetting.class);
                xlsSetting.putExtra("fileName", fileName);
                xlsSetting.putExtra("filePath", path.get(position));
                startActivity(xlsSetting);
            }
            else if(ext4Name.equals((String) "xlsx")){
                Toast.makeText(getApplicationContext(), "확장자를 xls 로 변경 후 사용해주세요.", Toast.LENGTH_SHORT).show();
            }
        }
    }
}
