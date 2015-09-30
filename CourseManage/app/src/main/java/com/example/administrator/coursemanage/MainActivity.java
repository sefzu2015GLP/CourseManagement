package com.example.administrator.coursemanage;

import java.io.FileInputStream;
import java.io.InputStream;

import android.os.Bundle;
import android.app.Activity;
import android.renderscript.ScriptIntrinsicYuvToRGB;
import android.text.method.ScrollingMovementMethod;
import android.view.Menu;
import android.widget.TextView;
import android.database.Cursor;
import android.database.sqlite.SQLiteDatabase;
import android.util.Log;
import android.content.ContentValues;

import jxl.*;

public class MainActivity extends Activity {
    TextView txt = null;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        txt = (TextView) findViewById(R.id.txt_show);
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
            String show = new String();//����һ���ַ��������ڳ�������ʱ��logcat�в鿴���
            txt.setText("the num of sheets is " + num + "\n");
            System.out.println("the num of sheet is +num +\n");
            // ��õ�һ�����������
            Sheet sheet = book.getSheet(0);
            int Rows = sheet.getRows();
            int Cols = sheet.getColumns();
            txt.append("the name of sheet is " + sheet.getName() + "\n");//append����Ϊ����ѡԪ�غ������Ӧ����
            show = show + "the name of sheet is " + sheet.getName() + "\n";
            txt.append("total rows is " + Rows + "\n");
            show = show + "total rows is " + Rows + "\n";
            txt.append("total cols is " + Cols + "\n");
            show = show + "total cols is " + Cols + "\n";
            for (int i = 0; i < Cols; ++i) {
                for (int j = 0; j < Rows; ++j) {
                    // getCell(Col,Row)��õ�Ԫ���ֵ
                    txt.append("contents:" + sheet.getCell(i, j).getContents() +",");
                    show = show + "contents:" + sheet.getCell(i, j).getContents()+"," ;

                }
                txt.append("\n");
                show = show + "\n";
            }
            System.out.println(show);
            book.close();
            //ֱ�Ӵ��ֻ���ָ��·�������ݿ�
            SQLiteDatabase db = openOrCreateDatabase("/storage/emulated/0/Tencent/QQfile_recv/new.s3db", MODE_PRIVATE, null);
            //�����½����ݿ�
            /*db.execSQL("create table Course(_id integer primary key,g text not null,major text not null,num text not null," +
                    "name text not null,choose text not null,mark text not null,texttime text not null," +
                    "atime text,btime text,week text,teacher text,ps text )");*/
            //��excel������ĵ�Ԫ�����ݲ������ݿ�
            for (int i = 0; i < Cols; ++i) {
                int j = 0;
                ContentValues values = new ContentValues();
                String content = sheet.getCell(i, j).getContents();
                values.put("�꼶", content);
                if (j < Rows) j++;
                content = sheet.getCell(i, j).getContents();
                values.put("רҵ", content);
                if (j < Rows) j++;
                content = sheet.getCell(i, j).getContents();
                values.put("רҵ����", content);
                if (j < Rows) j++;
                content = sheet.getCell(i, j).getContents();
                values.put("�γ�����", content);
                if (j < Rows) j++;
                content = sheet.getCell(i, j).getContents();
                values.put("ѡ������", content);
                if (j < Rows) j++;
                content = sheet.getCell(i, j).getContents();
                values.put("ѧ��", content);
                if (j < Rows) j++;
                content = sheet.getCell(i, j).getContents();
                values.put("ѧʱ", content);
                if (j < Rows) j++;
                content = sheet.getCell(i, j).getContents();
                values.put("ʵ��ѧʱ", content);
                if (j < Rows) j++;
                content = sheet.getCell(i, j).getContents();
                values.put("�ϻ�ѧʱ", content);
                if (j < Rows) j++;
                content = sheet.getCell(i, j).getContents();
                values.put("��������", content);
                if (j < Rows) j++;
                content = sheet.getCell(i, j).getContents();
                values.put("�ον�ʦ", content);
                if (j < Rows) j++;
                content = sheet.getCell(i, j).getContents();
                values.put("��ע", content);
                db.insert("course", null, values);
                values.clear();//ÿ����ɲ����values���
            }
            //��ѯ��������α�c�У�������logcat����ʾ
            Cursor c = db.rawQuery("select * from course", null);
            if (c != null) {
                while (c.moveToNext()) {
                    Log.i("info", "�꼶:" + c.getString(c.getColumnIndex("�꼶")));
                    Log.i("info", "רҵ:" + c.getString(c.getColumnIndex("רҵ")));
                    Log.i("info", "רҵ����:" + c.getString(c.getColumnIndex("רҵ����")));
                    Log.i("info", "�γ�����:" + c.getString(c.getColumnIndex("�γ�����")));
                    Log.i("info", "ѡ������:" + c.getString(c.getColumnIndex("ѡ������")));
                    Log.i("info", "ѧ��:" + c.getString(c.getColumnIndex("ѧ��")));
                    Log.i("info", "ѧʱ:" + c.getString(c.getColumnIndex("ѧʱ")));
                    Log.i("info", "ʵ��ѧʱ:" + c.getString(c.getColumnIndex("ʵ��ѧʱ")));
                    Log.i("info", "�ϻ�ѧʱ:" + c.getString(c.getColumnIndex("�ϻ�ѧʱ")));
                    Log.i("info", "��������:" + c.getString(c.getColumnIndex("��������")));
                    Log.i("info", "�ον�ʦ:" + c.getString(c.getColumnIndex("�ον�ʦ")));
                    Log.i("info", "��ע:" + c.getString(c.getColumnIndex("��ע")));
                    Log.i("info", "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                }
                c.close();
            }
            db.close();
        } catch (Exception e) {
            System.out.println(e);
        }
    }
}

