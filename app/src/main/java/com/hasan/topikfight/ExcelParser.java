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
import java.util.HashMap;
import java.util.LinkedList;
import java.util.Map;

class WordPair {
    String eng;
    String kor;
    public WordPair(String _eng, String _kor) {
        this.eng = _eng;
        this.kor = _kor;

    }

}

public class ExcelParser {

    final int SIZE = 100;
    LinkedList<WordPair> wordList[];
    String filename;

    int curRow;
    int curCol;

    int listSize = 0;
    final String TAG = "ExcelParse";

    ExcelParser(String _filename) {

        this.filename = _filename;
        wordList = new LinkedList[SIZE];
        for (int i=0; i<SIZE; i++)
            wordList[i] = new LinkedList<WordPair>();
    }

    LinkedList<WordPair> WordList(int index) {
        return wordList[index];
    }
    int ListSize() {
        return listSize;
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
            int index = -1;
            LinkedList<Integer> wordCols = new LinkedList<Integer>();
            Map<Integer, Integer> map = new HashMap<Integer, Integer>();
            int currPos = 0;
            WordPair curWordPair = new WordPair("", "");

            int prev = 0;
            String lastValue = "last";
            for (int r = 0; r<rowsCount; r++) {

                this.curRow = r;
                Row row = mySheet.getRow(r);
                int cellsCount = row.getPhysicalNumberOfCells();





                for (int c = 0; c<cellsCount; c++) {
                    this.curCol = c;
                    String value = getCellAsString(row, c, formulaEvaluator);
                    String cellInfo = "r:"+r+"; c:"+c+"; v:"+value;
                    //printlnToUser(cellInfo);


                    int targetCol = -1;
                    if (r > 0) {
                        targetCol = wordCols.get(currPos);
                        if (prev != targetCol) {
                            Log.e(TAG, "target col is: " + targetCol + " cur pos is: " + currPos + " cur row is: " + r + " cur col is: " + c);

                        }

                        prev = targetCol;
                    }

                    if (c == cellsCount-1)
                        lastValue = value;




                    if (r > 0 && c < targetCol-1 && c > targetCol+1)
                        continue;



                    float f;
                    try {
                        f = Float.parseFloat(value);

                    }
                    catch (Exception e) {
                        f = -1;
                    }



                    if (r == 0) {
                        final String WORD = "단어";
                        if (value.substring(0, WORD.length()).compareTo(WORD) == 0) {
                            wordCols.add(c);

                        }
                    }
                    else {
                        if (c == targetCol-1) {

                            index = (int) f;
                            if (index >= 1) {
                                map.put(targetCol, index - 1);
                                Log.e(TAG, "map values are: " + targetCol + " : " + (index-1));
                            }
                        }
                        else if (c == targetCol) {
                            curWordPair.kor = value;
                        }
                        else if (c == targetCol+1) {
                            //Log.e(TAG, "here 1");
                            curWordPair.eng = value;
                            int k = map.get(targetCol);
                            if (k < 0) {
                                Log.e(TAG, "Excel file format error: index found in the map is negative");
                                break;
                            }
                            else if (k > SIZE) {
                                Log.e(TAG, "Excel file format error: index found in the map is too large");
                                break;
                            }
                            if (k > listSize-1)
                                listSize = k+1;


                            wordList[k].add(new WordPair(curWordPair.eng, curWordPair.kor));



                            if (wordCols.size() > 0) {
                                if (currPos == wordCols.size()-1 ) {
                                    ;//Log.e(TAG, "NEXT");
                                }
                                currPos = (currPos + 1) % wordCols.size();
                                //Log.e(TAG, "inc done");

                            }
                            else {
                                Log.e(TAG, "Excel file format error: wordCols size is zero");
                                break;
                            }

                        }
                    }


                    if (r >= 1) {
                        //Log.e(TAG, "list is " + wordCols);
                        if (wordCols.size() == 0) {
                            Log.e(TAG, "Excel file format error : list is empty");
                            break;
                        }
                    }




                }

                if (r < 50) {
                    Log.e(TAG, "index: " + r + " no of rows " + rowsCount + " no of cells: " + cellsCount + " last value is: " + lastValue);
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
            if (cellValue == null)
                return "NULL";

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
