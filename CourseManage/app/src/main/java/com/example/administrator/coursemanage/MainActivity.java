package com.example.administrator.coursemanage;

import java.io.FileInputStream;
import java.io.InputStream;

import android.os.Bundle;
import android.app.Activity;
import android.text.method.ScrollingMovementMethod;
import android.view.Menu;
import android.widget.TextView;

import jxl.*;

public class MainActivity extends Activity {
    TextView txt = null;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        txt = (TextView)findViewById(R.id.txt_show);
        txt.setMovementMethod(ScrollingMovementMethod.getInstance());
        readExcel();
    }

    @Override
    public boolean onCreateOptionsMenu(Menu menu) {
        // Inflate the menu; this adds items to the action bar if it is present.
        getMenuInflater().inflate(R.menu.menu_main, menu);
        return true;
    }

    public void readExcel() {
        try {
            InputStream is = new FileInputStream("/storage/emulated/0/Tencent/QQfile_recv/data.xls");//��ȡ�ֻ��ڴ���ָ��·�����ļ�
            //Workbook book = Workbook.getWorkbook(new File("mnt/sdcard/test.xls"));
            Workbook book = Workbook.getWorkbook(is);
            int num = book.getNumberOfSheets();//��ȡsheet����Ŀ
            String show=new String();//����һ���ַ��������ڳ�������ʱ��logcat�в鿴���
            txt.setText("the num of sheets is " + num+ "\n");
            System.out.println("the num of sheet is +num +\n");
            // ��õ�һ�����������
            Sheet sheet = book.getSheet(0);
            int Rows = sheet.getRows();
            int Cols = sheet.getColumns();
            txt.append("the name of sheet is " + sheet.getName() + "\n");//append����Ϊ����ѡԪ�غ������Ӧ����
            show=show+"the name of sheet is " + sheet.getName() + "\n";
            txt.append("total rows is " + Rows + "\n");
            show=show+"total rows is " + Rows + "\n";
            txt.append("total cols is " + Cols + "\n");
            show=show+"total cols is " + Cols + "\n";
            for (int i = 0; i < Cols; ++i) {
                for (int j = 0; j < Rows; ++j) {
                    // getCell(Col,Row)��õ�Ԫ���ֵ
                    txt.append("contents:" + sheet.getCell(i,j).getContents() + "\n");
                    show=show+"contents:" + sheet.getCell(i,j).getContents() + "\n";

                }
                show=show+"\n";
            }
            System.out.println(show);
            book.close();
        } catch (Exception e) {
            System.out.println(e);
        }
    }

}