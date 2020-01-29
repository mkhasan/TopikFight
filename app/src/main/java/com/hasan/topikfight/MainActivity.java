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
import android.view.View;
import android.widget.Button;
import android.widget.TextView;

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


public class MainActivity extends AppCompatActivity {

    TextView txtView;
    Button btnRead;

    final String TAG = "MainActivity";

    LinkedList<WordPair> wordList;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        System.setProperty("org.apache.poi.javax.xml.stream.XMLInputFactory", "com.fasterxml.aalto.stax.InputFactoryImpl");
        System.setProperty("org.apache.poi.javax.xml.stream.XMLOutputFactory", "com.fasterxml.aalto.stax.OutputFactoryImpl");
        System.setProperty("org.apache.poi.javax.xml.stream.XMLEventFactory", "com.fasterxml.aalto.stax.EventFactoryImpl");

        txtView = findViewById(R.id.textView);
        btnRead = findViewById(R.id.button);

        if (!checkPermissionForReadExtertalStorage()) {
            try {
                requestPermissionForReadExtertalStorage();
            } catch (Exception e) {
                e.printStackTrace();
                Log.e("eRROR", "ERROR");
            }
        }

        btnRead.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                //readExcelFileFromAssets();

                ExcelParser excelParser = new ExcelParser(Environment.getExternalStorageDirectory()+"/Documents/Chuk_Chuk_TOPIK_Lists.xlsx");
                excelParser.Parse();

                LinkedList<WordPair> currList = excelParser.WordList(5);

                String msg = "";

                Iterator<WordPair> it = currList.iterator();
                while(it.hasNext()) {
                    WordPair wordPair = it.next();
                    msg += wordPair.eng + ":" + wordPair.kor + "\n";
                }

                Log.e(TAG, msg);


            }
        });




    }

    public void readExcelFileFromAssets() {
        try {
            /*
            InputStream myInput;
            // initialize asset manager
            AssetManager assetManager = getAssets();
            //  open excel sheet
            myInput = assetManager.open("myexcelsheet.xls");
            // Create a POI File System object
            POIFSFileSystem myFileSystem = new POIFSFileSystem(myInput);
            // Create a workbook using the File System
            HSSFWorkbook myWorkBook = new HSSFWorkbook(myFileSystem);
            // Get the first sheet from workbook
            HSSFSheet mySheet = myWorkBook.getSheetAt(0);
            // We now need something to iterate through the cells.
            Iterator<Row> rowIter = mySheet.rowIterator();
            int rowno =0;
            txtView.append("\n");
            while (rowIter.hasNext()) {
                Log.e("aaa", " row no "+ rowno );
                HSSFRow myRow = (HSSFRow) rowIter.next();
                if(rowno !=0) {
                    Iterator<Cell> cellIter = myRow.cellIterator();
                    int colNum =0;
                    String sno="", date="", det="";
                    while (cellIter.hasNext()) {
                        HSSFCell myCell = (HSSFCell) cellIter.next();
                        if (colNum==0){
                            sno = myCell.toString();
                        }else if (colNum==1){
                            date = myCell.toString();
                        }else if (colNum==2){
                            det = myCell.toString();
                        }
                        colNum++;
                        Log.e("aaa", " Index :" + myCell.getColumnIndex() + " -- " + myCell.toString());
                    }
                    txtView.append( sno + " -- "+ date+ "  -- "+ det+"\n");
                }
                rowno++;


            }
            */

            File myFile = new File(Environment.getExternalStorageDirectory()+"/Documents/Chuk_Chuk_TOPIK_-Lists.xlsx");


            FileInputStream fis = new FileInputStream(myFile);

// Finds the workbook instance for XLSX file
            XSSFWorkbook myWorkBook = new XSSFWorkbook (fis);
// Return first sheet from the XLSX workbook
            XSSFSheet mySheet = myWorkBook.getSheetAt(0);


            int rowsCount = mySheet.getPhysicalNumberOfRows();
            Log.e("Rows", "rows count " + rowsCount);

            FormulaEvaluator formulaEvaluator = myWorkBook.getCreationHelper().createFormulaEvaluator();
            for (int r = 0; r<rowsCount; r++) {
                Row row = mySheet.getRow(r);
                int cellsCount = row.getPhysicalNumberOfCells();
                if (r < 0)
                    Log.e("Cols", "ros pos " + r + " cols count" + cellsCount);

                for (int c = 0; c<cellsCount; c++) {
                    String value = getCellAsString(row, c, formulaEvaluator);
                    String cellInfo = "r:"+r+"; c:"+c+"; v:"+value;
                    //printlnToUser(cellInfo);
                    if (r < 10)
                        Log.e("TAG", "rows " + r + " cell " + c + " cellInfo " + cellInfo);
                }


            }



// Get iterator to all the rows in current sheet
            /*
            Iterator rowIterator = mySheet.iterator();
// Traversing over each row of XLSX file
            int rowsCount = mySheet.getPhysicalNumberOfRows();
            for (int r=0; r<rowsCount; r++)
            {
                Row row =
// For each row, iterate through each columns
                Iterator cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    switch (cell.getCellType()) { case Cell.CELL_TYPE_STRING: System.out.print(cell.getStringCellValue() + “\t”);
                        break; case Cell.CELL_TYPE_NUMERIC:System.out.print(cell.getNumericCellValue() + “\t”);
                        break; case Cell.CELL_TYPE_BOOLEAN: System.out.print(cell.getBooleanCellValue() + “\t”);
                        break; default : }
                }
                System.out.println(“”);
            }


             */

        } catch (Exception e) {


            Log.e("aaa", "error " + Environment.getExternalStorageDirectory()+ " " + e.toString());
        }
    }
    private void printlnToUser(String str) {
        final String string = str;

        Log.d("Error", "someOtherMethod()");



    }

    protected String getCellAsString(Row row, int c, FormulaEvaluator formulaEvaluator) {
        String value = "";
        try {
            Cell cell = row.getCell(c);
            CellValue cellValue = formulaEvaluator.evaluate(cell);
            switch (cellValue.getCellType()) {
                case Cell.CELL_TYPE_BOOLEAN:
                    value = ""+cellValue.getBooleanValue();
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    double numericValue = cellValue.getNumberValue();
                    if(HSSFDateUtil.isCellDateFormatted(cell)) {
                        double date = cellValue.getNumberValue();
                        SimpleDateFormat formatter =
                                new SimpleDateFormat("dd/MM/yy");
                        value = formatter.format(HSSFDateUtil.getJavaDate(date));
                    } else {
                        value = ""+numericValue;
                    }
                    break;
                case Cell.CELL_TYPE_STRING:
                    value = ""+cellValue.getStringValue();
                    break;
                default:
            }
        } catch (NullPointerException e) {
            /* proper error handling should be here */
            printlnToUser(e.toString());
        }
        return value;
    }

    public String getDataDir(final Context context) throws Exception {
        return context.getPackageManager().getPackageInfo(context.getPackageName(), 0).applicationInfo.dataDir;
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
}
