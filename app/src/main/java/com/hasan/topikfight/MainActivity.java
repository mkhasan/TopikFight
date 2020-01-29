package com.hasan.topikfight;

import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;


import android.Manifest;
import android.app.Activity;
import android.content.Context;
import android.content.ContextWrapper;
import android.content.pm.PackageManager;
import android.content.res.AssetManager;

import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.util.Log;
import android.util.TypedValue;
import android.view.MotionEvent;
import android.view.View;
import android.widget.AdapterView;
import android.widget.ArrayAdapter;
import android.widget.Button;
import android.widget.Spinner;
import android.widget.TextView;
import android.widget.Toast;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;



import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Iterator;
import java.util.LinkedList;

import static java.lang.System.exit;


public class MainActivity extends AppCompatActivity implements AdapterView.OnItemSelectedListener {

    //TextView txtView;
    Button btnRead;

    final String TAG = "MainActivity";

    LinkedList<WordPair> wordList;

    String listSelector[];

    int curSel = 0;
    int wordIndex = 0;
    int listSize;
    boolean showKorean;

    ExcelParser excelParser;
    TextView wordView;
    TextView wordIndexView;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        System.setProperty("org.apache.poi.javax.xml.stream.XMLInputFactory", "com.fasterxml.aalto.stax.InputFactoryImpl");
        System.setProperty("org.apache.poi.javax.xml.stream.XMLOutputFactory", "com.fasterxml.aalto.stax.OutputFactoryImpl");
        System.setProperty("org.apache.poi.javax.xml.stream.XMLEventFactory", "com.fasterxml.aalto.stax.EventFactoryImpl");

        //txtView = findViewById(R.id.textView);
        wordView = findViewById(R.id.word);
        wordIndexView = findViewById(R.id.word_index);

        excelParser = new ExcelParser(Environment.getExternalStorageDirectory()+"/Documents/Chuk_Chuk_TOPIK_Lists.xlsx");
        excelParser.Parse();

        listSize = excelParser.ListSize();
        if (listSize < 1) {
            exit(1);

        }

        listSelector = new String[listSize];
        for (int k=0; k<listSize; k++) {
            listSelector[k] = "Word List: " + (k+1);
        }




        btnRead = findViewById(R.id.button);

        if (!checkPermissionForReadExtertalStorage()) {
            try {
                requestPermissionForReadExtertalStorage();
            } catch (Exception e) {
                e.printStackTrace();
                Log.e("eRROR", "ERROR");
            }
        }



        Spinner spinner = (Spinner) findViewById(R.id.list_spinner);
// Create an ArrayAdapter using the string array and a default spinner layout
        //ArrayAdapter<CharSequence> adapter = ArrayAdapter.createFromResource(this,
          //      R.array.planets_array, android.R.layout.simple_spinner_item);
        ArrayAdapter<String> adapter = new ArrayAdapter<String>(this, android.R.layout.simple_spinner_item, listSelector);

// Specify the layout to use when the list of choices appears
        adapter.setDropDownViewResource(android.R.layout.simple_spinner_dropdown_item);
// Apply the adapter to the spinner
        spinner.setAdapter(adapter);
        spinner.setOnItemSelectedListener(this);

        btnRead.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                //readExcelFileFromAssets();

                Fetch();

            }
        });

        findViewById(R.id.flip).setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                if (showKorean)
                    showKorean = false;
                else
                    showKorean = true;
                UpdateView();
            }
        });

        wordView.setOnTouchListener(new OnSwipeTouchListener() {
            public void onSwipeTop() {
                //Toast.makeText(MainActivity.this, "top", Toast.LENGTH_SHORT).show();
            }
            public void onSwipeRight() {
                //Toast.makeText(MainActivity.this, "right", Toast.LENGTH_SHORT).show();
                int k = excelParser.WordList(curSel).size();
                wordIndex = (wordIndex+k-1) % k;
                showKorean = true;
                UpdateView();

            }
            public void onSwipeLeft() {
                //Toast.makeText(MainActivity.this, "left", Toast.LENGTH_SHORT).show();
                int k = excelParser.WordList(curSel).size();
                wordIndex = (wordIndex+1) % k;
                showKorean = true;
                UpdateView();
            }
            public void onSwipeBottom() {
                //Toast.makeText(MainActivity.this, "bottom", Toast.LENGTH_SHORT).show();

            }

            public boolean onTouch(View v, MotionEvent event) {
                return gestureDetector.onTouchEvent(event);
            }

            public void doubleTapeHandler() {
                Toast.makeText(MainActivity.this, "Reset", Toast.LENGTH_SHORT).show();
                wordIndex = 0;
                showKorean = true;
                UpdateView();
            }

        });




    }

    public void onItemSelected(AdapterView<?> parent, View view,
                               int pos, long id) {
        // An item was selected. You can retrieve the selected item using
        // parent.getItemAtPosition(pos)
        Log.e(TAG, "Selected " + id + " pos is " + pos);
        curSel = pos;
        wordIndex = 0;
        showKorean = true;
        UpdateView();
    }

    public void onNothingSelected(AdapterView<?> parent) {
        // Another interface callback
        Log.e(TAG, "Nothing");
    }

    public void UpdateView() {
        LinkedList<WordPair> list = excelParser.WordList(curSel);
        //String kor = list[wordIndex].

        WordPair wordPair = list.get(wordIndex);

        if (wordPair == null) {
            Log.e(TAG, "Format error");
            exit(1);
        }

        wordIndexView.setText(Integer.toString(wordIndex+1));

        if (showKorean) {
            wordView.setTextSize(TypedValue.COMPLEX_UNIT_SP,75);
            wordView.setText(wordPair.kor);
        }
        else {
            wordView.setTextSize(TypedValue.COMPLEX_UNIT_SP,25);
            wordView.setText(wordPair.eng);
        }

        //msg += wordPair.eng + ":" + wordPair.kor + "\n";
    }

    public boolean checkPermissionForReadExtertalStorage() {
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.M) {
            int result = checkSelfPermission(Manifest.permission.READ_EXTERNAL_STORAGE);
            return result == PackageManager.PERMISSION_GRANTED;
        }
        return false;
    }
    public void requestPermissionForReadExtertalStorage() throws Exception {
        try {
            ActivityCompat.requestPermissions(this, new String[]{Manifest.permission.READ_EXTERNAL_STORAGE}, 3);
        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        }
    }

    public void Fetch() {
        ExcelParser excelParser = new ExcelParser(Environment.getExternalStorageDirectory()+"/Documents/Chuk_Chuk_TOPIK_Lists.xlsx");
        excelParser.Parse();

        LinkedList<WordPair> currList = excelParser.WordList(5);

        Log.e(TAG, "size of list is: " + excelParser.ListSize());
        String msg = "";

        Iterator<WordPair> it = currList.iterator();
        while(it.hasNext()) {
            WordPair wordPair = it.next();
            msg += wordPair.eng + ":" + wordPair.kor + "\n";
        }

        Log.e(TAG, msg);

    }
}
