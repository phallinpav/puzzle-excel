import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class Solution2 {
    private static final String FILE_NAME = "puzzle.xlsx";
    private static XSSFWorkbook workbook;
    private static XSSFSheet sheet;
    private static boolean done = false;
    private static XSSFCellStyle styleCorrectPath;

    static {
        File file = new File(FILE_NAME);
        FileInputStream fIP;
        try {
            fIP = new FileInputStream(file);
            //Get the workbook instance for XLSX file
            workbook = new XSSFWorkbook(fIP);
            sheet = workbook.getSheetAt(0);
            styleCorrectPath = workbook.createCellStyle();
            styleCorrectPath.setFillForegroundColor(IndexedColors.BLUE.getIndex());
            styleCorrectPath.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException(e);
        }
    }

    public static void start() {
        // ---------STARTING---------
        System.out.println("Start solving...");
        long startTime = System.currentTimeMillis();

        /* thread to show process loading...
        Thread thread = new Thread(new Runnable() {
            @Override
            public void run() {
                try {
                    Thread.sleep(1000);
                    int i = 1;
                    while (!done) {
                        System.out.print(".");
                        if (i++ == 3) {
                            System.out.print("\r");
                        }
                        Thread.sleep(400);
                    }
                } catch (Exception e) {
                    throw new RuntimeException(e);
                }
            }
        });
        thread.start();
        */

        // --------- ########## ---------
        // ---------SOLUTION CODE---------
        boolean isSolve = solvingMaze();
        if (isSolve) {
            try {
                saveFile();
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        } else {
            System.out.println("Unsolvable maze!!!!");
            return;
        }

        // --------- ########## ---------
        // ---------FINISHED---------
        long endTime = System.currentTimeMillis();
        long duration = endTime - startTime;
        System.out.printf("\r%nPuzzle is solved in %,d millisecond = %.2f second", duration, duration / 1000d);
        // ---------########---------
    }

    private static boolean solvingMaze() {
        // WRITE PROCESS HERE...

        boolean finish = checkAround(1,1);
        getCellAt(1, 1).setCellValue("S");
        done = true;
//
//        XSSFCell cell2 = getCellAt(4, 2);
//        cell2.setCellType(CellType.FORMULA);
//        cell2.setCellFormula("SUM(B5, 1)");
//
//        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
//        evaluator.evaluateAll();
//
//        CellValue cellValue = evaluator.evaluate(cell2);
//        double calculatedValue = cellValue.getNumberValue();
//        System.out.println("Calculated value: " + calculatedValue);
//
//        // Alternatively, you can use getNumericCellValue()
//        double numericValue = cell2.getNumericCellValue();
//        System.out.println("Numeric value: " + numericValue);
        return true;
    }

    private static boolean checkAround(int rowNum, int colNum) {
        if (isFinish(rowNum, colNum)) {
            return true;
        } else if (isVisited(rowNum, colNum)){
            return false;
        } else {
            XSSFCell cell = getCellAt(rowNum, colNum);
            cell.setCellValue("O");
            cell.setCellStyle(styleCorrectPath);
        }
        boolean check;
        // up
        if (isOpen(rowNum - 1, colNum)) {
            check = checkAround(rowNum - 1, colNum);
            if (check) {
                return true;
            }
        }
        // right
        if (isOpen(rowNum, colNum + 1)) {
            check = checkAround(rowNum, colNum + 1);
            if (check) {
                return true;
            }
        }
        //down
        if (isOpen(rowNum + 1, colNum)) {
            check = checkAround(rowNum + 1, colNum);
            if (check) {
                return true;
            }
        }
        // left
        if (isOpen(rowNum, colNum - 1)) {
            check = checkAround(rowNum, colNum - 1);
            if (check) {
                return true;
            }
        }
        getCellAt(rowNum, colNum).setCellValue("-");
        return false;
    }

    private static boolean isOpen(int rowNum, int colNum) {
        return !getCellAt(rowNum, colNum).getStringCellValue().equals("X");
    }

    private static boolean isFinish(int rowNum, int colNum) {
        return getCellAt(rowNum, colNum).getStringCellValue().equals("E");
    }

    private static boolean isVisited(int rowNum, int colNum) {
        return getCellAt(rowNum, colNum).getStringCellValue().equals("O");
    }

    private static XSSFCell getCellAt(int rowNum, int colNum) {
        XSSFRow row = sheet.getRow(rowNum);
        if (row == null) {
            row = sheet.createRow(rowNum);
        }
        XSSFCell cell = row.getCell(colNum);
        if (cell == null) {
            cell = row.createCell(colNum);
        }
        return cell;
    }

    private static void saveFile() throws IOException {
        //Create file system using specific name
        FileOutputStream out = new FileOutputStream(FILE_NAME);

        //write operation workbook using file out object
        workbook.write(out);
        out.close();
    }
}
