import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class Solution {
    private static final String FILE_NAME = "puzzle.xlsx";
    private static XSSFWorkbook workbook;
    private static XSSFSheet sheet;
    private static boolean done = false;
    private static boolean isReached = false;

    static {
        File file = new File(FILE_NAME);
        FileInputStream fIP;
        try {
            fIP = new FileInputStream(file);
            //Get the workbook instance for XLSX file
            workbook = new XSSFWorkbook(fIP);
            sheet = workbook.getSheetAt(0);
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

//        boolean a = fillField(1, 1);
        boolean finish = checkAround(1,1);
        CellStyle lightGreen = createCellStyle(IndexedColors.BRIGHT_GREEN);
        getCellAt(1, 1).setCellValue("S");
        getCellAt(1, 1).setCellStyle(lightGreen);

        done = true;

        return true;
    }

    private static void printMaze() {

    }


    private static void printMaze1() {
        try {
            Thread.sleep(300);
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }
        int maxCol = 0;
        int row = 0;
        int col = 0;
        String valCell = getCellAt(row, col).getStringCellValue();
        StringBuilder print = new StringBuilder();
        while(valCell.equals("X")) {
            print.append("X\t");
            col++;
            valCell = getCellAt(row, col).getStringCellValue();
        }
        maxCol = col;
        col = 0;
        row++;
        print.append("\n");

        valCell = getCellAt(row, col).getStringCellValue();
        while (valCell.equals("X")) {
            while (col <= maxCol) {
                valCell = getCellAt(row, col).getStringCellValue();
                print.append(valCell).append("\t");
                col++;
            }
            print.append("\n");
            row++;
            col = 0;
            valCell = getCellAt(row, col).getStringCellValue();
        }

        System.out.println(print);
    }

    private static boolean checkAround(int row, int col) {
        CellStyle cellStyleOrange = createCellStyle(IndexedColors.ORANGE);
        CellStyle cellStyleSkyBlue = createCellStyle(IndexedColors.SKY_BLUE);
        CellStyle cellStyleRed = createCellStyle(IndexedColors.RED);

        if (!isReached) {
            if (isFinish(row, col)) {
                return true;
            } else if (isTaken(row, col)) {
                return false;
            } else {
                getCellAt(row, col).setCellValue("O");
                getCellAt(row, col).setCellStyle(cellStyleOrange);
                printMaze();
            }
        } else {
            if (isFilled(row, col)) {
                return false;
            } else {
                getCellAt(row, col).setCellValue("@");
                getCellAt(row, col).setCellStyle(cellStyleRed);
                printMaze();
            }
        }



        boolean check;

        if (isOpen(row, col - 1)) {
            check = checkAround(row, col - 1);
            if (check) {
                isReached = true;
            }
        }

        if (isOpen(row - 1, col)) {
            check = checkAround(row - 1, col);
            if (check) {
                isReached = true;
            }
        }

        if (isOpen(row, col + 1)) {
            check = checkAround(row, col + 1);
            if (check) {
                isReached = true;
            }
        }

        if (isOpen(row + 1, col)) {
            check = checkAround(row + 1, col);
            if (check) {
                isReached = true;
            }
        }

        if (!isReached) {
            getCellAt(row, col).setCellValue("-");
            getCellAt(row, col).setCellStyle(cellStyleSkyBlue);
            printMaze();
        }

        return false;
    }

    private static CellStyle createCellStyle(IndexedColors colors) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFillForegroundColor(colors.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return cellStyle;
    }


    private static boolean isOpen(int row, int col) {
        return !getCellAt(row, col).getStringCellValue().equals("X");
    }

    private static boolean isTaken(int row, int col) {
        return getCellAt(row, col).getStringCellValue().equals("O");
    }

    private static boolean isFilled(int row, int col) {
        return getCellAt(row, col).getStringCellValue().equals("@") || getCellAt(row, col).getStringCellValue().equals("O") || getCellAt(row, col).getStringCellValue().equals("-");
    }

    private static boolean isFinish(int row, int col) {
        return getCellAt(row, col).getStringCellValue().equals("E");
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
