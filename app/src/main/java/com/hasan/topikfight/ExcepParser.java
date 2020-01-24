package com.hasan.topikfight;

import android.os.Environment;
import android.util.Log;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.LinkedList;

class WordPair {
    String eng;
    String kor;
    private WordPair(String _eng, String _kor) {
        this.eng = _eng;
        this.kor = _kor;

    }

}

public class ExcepParser {

    LinkedList<WordPair> wordList;
    String filename;

    final String TAG = "ExcelParse";

    ExcepParser(String _filename) {
        this.filename = _filename;
        wordList = new LinkedList<WordPair>();
    }

    LinkedList<WordPair> WordList() {
        return wordList;
    }

    boolean Parse(){


        try {

            File myFile = new File(filename);


            FileInputStream fis = new FileInputStream(myFile);


            XSSFWorkbook myWorkBook = new XSSFWorkbook (fis);

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




        } catch (Exception e) {


            Log.e("aaa", "error " + Environment.getExternalStorageDirectory()+ " " + e.toString());
            return false;
        }



        return true;
    }


    protected String getCellAsString(Row row, int c, FormulaEvaluator formulaEvaluator) {
        String value = "";
        try {
            Cell cell = row.getCell(c);
            CellValue cellValue = formulaEvaluator.evaluate(cell);
            switch (cellValue.getCellType()) {
                case Cell.CELL_TYPE_BOOLEAN:
                    value = "" + cellValue.getBooleanValue();
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    double numericValue = cellValue.getNumberValue();
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        double date = cellValue.getNumberValue();
                        SimpleDateFormat formatter =
                                new SimpleDateFormat("dd/MM/yy");
                        value = formatter.format(HSSFDateUtil.getJavaDate(date));
                    } else {
                        value = "" + numericValue;
                    }
                    break;
                case Cell.CELL_TYPE_STRING:
                    value = "" + cellValue.getStringValue();
                    break;
                default:
            }
        } catch (NullPointerException e) {
            /* proper error handling should be here */
            Log.e(TAG, e.toString());
        }
        return value;
    }


}
