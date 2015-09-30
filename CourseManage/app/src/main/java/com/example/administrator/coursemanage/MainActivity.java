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
            InputStream is = new FileInputStream("/storage/emulated/0/Tencent/QQfile_recv/data.xls");//获取手机内存中指定路径的文件
            //Workbook book = Workbook.getWorkbook(new File("mnt/sdcard/test.xls"));
            Workbook book = Workbook.getWorkbook(is);
            int num = book.getNumberOfSheets();//获取sheet的数目
            String show = new String();//定义一个字符串用于在程序运行时在logcat中查看结果
            txt.setText("the num of sheets is " + num + "\n");
            System.out.println("the num of sheet is +num +\n");
            // 获得第一个工作表对象
            Sheet sheet = book.getSheet(0);
            int Rows = sheet.getRows();
            int Cols = sheet.getColumns();
            txt.append("the name of sheet is " + sheet.getName() + "\n");//append方法为在所选元素后添加相应参数
            show = show + "the name of sheet is " + sheet.getName() + "\n";
            txt.append("total rows is " + Rows + "\n");
            show = show + "total rows is " + Rows + "\n";
            txt.append("total cols is " + Cols + "\n");
            show = show + "total cols is " + Cols + "\n";
            for (int i = 0; i < Cols; ++i) {
                for (int j = 0; j < Rows; ++j) {
                    // getCell(Col,Row)获得单元格的值
                    txt.append("contents:" + sheet.getCell(i, j).getContents() +",");
                    show = show + "contents:" + sheet.getCell(i, j).getContents()+"," ;

                }
                txt.append("\n");
                show = show + "\n";
            }
            System.out.println(show);
            book.close();
            //直接打开手机中指定路径的数据库
            SQLiteDatabase db = openOrCreateDatabase("/storage/emulated/0/Tencent/QQfile_recv/new.s3db", MODE_PRIVATE, null);
            //测试新建数据库
            /*db.execSQL("create table Course(_id integer primary key,g text not null,major text not null,num text not null," +
                    "name text not null,choose text not null,mark text not null,texttime text not null," +
                    "atime text,btime text,week text,teacher text,ps text )");*/
            //将excel解析后的单元格内容插入数据库
            for (int i = 0; i < Cols; ++i) {
                int j = 0;
                ContentValues values = new ContentValues();
                String content = sheet.getCell(i, j).getContents();
                values.put("年级", content);
                if (j < Rows) j++;
                content = sheet.getCell(i, j).getContents();
                values.put("专业", content);
                if (j < Rows) j++;
                content = sheet.getCell(i, j).getContents();
                values.put("专业人数", content);
                if (j < Rows) j++;
                content = sheet.getCell(i, j).getContents();
                values.put("课程名称", content);
                if (j < Rows) j++;
                content = sheet.getCell(i, j).getContents();
                values.put("选修类型", content);
                if (j < Rows) j++;
                content = sheet.getCell(i, j).getContents();
                values.put("学分", content);
                if (j < Rows) j++;
                content = sheet.getCell(i, j).getContents();
                values.put("学时", content);
                if (j < Rows) j++;
                content = sheet.getCell(i, j).getContents();
                values.put("实验学时", content);
                if (j < Rows) j++;
                content = sheet.getCell(i, j).getContents();
                values.put("上机学时", content);
                if (j < Rows) j++;
                content = sheet.getCell(i, j).getContents();
                values.put("起讫周序", content);
                if (j < Rows) j++;
                content = sheet.getCell(i, j).getContents();
                values.put("任课教师", content);
                if (j < Rows) j++;
                content = sheet.getCell(i, j).getContents();
                values.put("备注", content);
                db.insert("course", null, values);
                values.clear();//每次完成插入后将values清空
            }
            //查询结果放入游标c中，遍历在logcat中显示
            Cursor c = db.rawQuery("select * from course", null);
            if (c != null) {
                while (c.moveToNext()) {
                    Log.i("info", "年级:" + c.getString(c.getColumnIndex("年级")));
                    Log.i("info", "专业:" + c.getString(c.getColumnIndex("专业")));
                    Log.i("info", "专业人数:" + c.getString(c.getColumnIndex("专业人数")));
                    Log.i("info", "课程名称:" + c.getString(c.getColumnIndex("课程名称")));
                    Log.i("info", "选修类型:" + c.getString(c.getColumnIndex("选修类型")));
                    Log.i("info", "学分:" + c.getString(c.getColumnIndex("学分")));
                    Log.i("info", "学时:" + c.getString(c.getColumnIndex("学时")));
                    Log.i("info", "实验学时:" + c.getString(c.getColumnIndex("实验学时")));
                    Log.i("info", "上机学时:" + c.getString(c.getColumnIndex("上机学时")));
                    Log.i("info", "起讫周序:" + c.getString(c.getColumnIndex("起讫周序")));
                    Log.i("info", "任课教师:" + c.getString(c.getColumnIndex("任课教师")));
                    Log.i("info", "备注:" + c.getString(c.getColumnIndex("备注")));
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

