import org.apache.poi.ss.usermodel.FillPatternType;
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
import java.util.ArrayDeque;
import java.util.Queue;

public class Solution4 {
    private static final String FILE_NAME = "puzzle.xlsx";
    private static final XSSFWorkbook workbook;
    private static final XSSFSheet sheet;
    private static boolean done = false;
    private static final XSSFCellStyle styleCorrectPath;
    private static final XSSFCellStyle styleInCorrectPath;

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
            styleInCorrectPath = workbook.createCellStyle();
            styleInCorrectPath.setFillForegroundColor(IndexedColors.RED.getIndex());
            styleInCorrectPath.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException(e);
        }
    }

    public static void start() {
        // ---------STARTING---------
        System.out.println("Start solving...");
        long startTime = System.currentTimeMillis();

        // --------- ########## ---------
        // ---------SOLUTION CODE---------
        boolean isSolve = solvingMaze();
        if (isSolve) {
//            try {
//                saveFile();
//            } catch (IOException e) {
//                throw new RuntimeException(e);
//            }
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

    private static void printMaze() {}

    private static void printMaze1() {
//        try {
//            Thread.sleep(300);
//        } catch (InterruptedException e) {
//            throw new RuntimeException(e);
//        }
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

        System.out.println("=======================================================");
        System.out.println(print);
        System.out.println("=======================================================");
    }

    private static boolean solvingMaze() {
        // WRITE PROCESS HERE...

        solution();

//        done = true;
        return done;
    }

    private static final Queue<Integer[]> queue1 = new ArrayDeque<>();
    private static final Queue<Integer[]> queue2 = new ArrayDeque<>();
    private static Integer[] lastMove = new Integer[]{};
    private static int count = 1;

    private static void solution() {
        addPossibleMoves(1, 1);
        queue1.addAll(queue2);
        queue2.clear();
        while (!queue1.isEmpty()) {
            Integer[] move = queue1.poll();
            if (checkCurrentMove(move[0], move[1])) {
                addPossibleMoves(move[0], move[1]);
                XSSFCell cell = getCellAt(move[0], move[1]);
                cell.setCellValue(String.valueOf(count));
                printMaze();
            }
            if (queue1.isEmpty()) {
                System.out.println("-----------------------");
                System.out.println("COUNT = " + count);
                System.out.println("NEXT MOVES = " + queue2.size());
                System.out.println("-----------------------");
                queue1.addAll(queue2);
                queue2.clear();
                count++;
            }
        }
        addColorToShortestPath();
    }

    private static void addColorToShortestPath() {
        if (done) {
            Integer[] temp = lastMove;
            int i = 0;
            int j = -1;
            count--;
            while (count > 0) {
                XSSFCell cell = getCellAt(temp[0] + i, temp[1] + j);
                if (cell.getStringCellValue().equals(String.valueOf(count))) {
                    cell.setCellStyle(styleCorrectPath);
                    temp = new Integer[] {temp[0] + i, temp[1] + j};
                    i = 0;
                    j = -1;
                    System.out.println("COUNT BACK = " + count);
                    count--;
                } else {
                    if (j == -1) {
                        i = -1;
                        j = 0;
                    } else if (i == -1) {
                        i = 0;
                        j = 1;
                    } else if (j == 1) {
                        i = 1;
                        j = 0;
                    }
                }
            }
        }
    }

    private static boolean checkCurrentMove(int row, int col) {
        if (isFinish(row, col)) {
            queue1.clear();
            queue2.clear();
            lastMove = new Integer[] { row, col };
            done = true;
            return false;
        } else if (isVisited(row, col)) {
            return false;
        } else if (isOpen(row, col)) {
            return true;
        }
        return false;
    }

    private static void addPossibleMoves(int row, int col) {
        Integer[] right = new Integer[]{row, col + 1};
        Integer[] down = new Integer[]{row + 1, col};
        Integer[] left = new Integer[]{row, col - 1};
        Integer[] up = new Integer[]{row - 1, col};
        addPossibleMove(right);
        addPossibleMove(down);
        addPossibleMove(left);
        addPossibleMove(up);
    }

    private static void addPossibleMove(Integer[] move) {
        if (checkCurrentMove(move[0], move[1])) {
            queue2.add(move);
        }
    }

    // ===================================================
    // COMMON METHOD TO CHECK CELL MAZE
    // ===================================================


    // ===================================================
    // NOT EVEN USE ANYMORE FOR THESE
    // ===================================================
    private static boolean isOpen(int rowNum, int colNum) {
        return isOpen(getCellAt(rowNum, colNum));
    }

    private static boolean isOpen(XSSFCell cell) {
        return isOpen(getCellStringValueAt(cell));
    }

    private static boolean isOpen(String st) {
        return !(st.equals("S") || st.equals("X"));
    }

    private static boolean isFinish(int rowNum, int colNum) {
        return isFinish(getCellAt(rowNum, colNum));
    }

    private static boolean isFinish(XSSFCell cell) {
        return isFinish(getCellStringValueAt(cell));
    }

    private static boolean isFinish(String st) {
        return st.equals("E");
    }

    private static boolean isVisited(int rowNum, int colNum) {
        XSSFCell cell = getCellAt(rowNum, colNum);
        return isVisited(cell);
    }

    private static boolean isVisited(XSSFCell cell) {
        String st = getCellStringValueAt(cell);
        return isVisited(st);
    }

    private static boolean isVisited(String st) {
        try {
            int i = Integer.parseInt(st);
            return i > 0;
        } catch (NumberFormatException ex) {
            return false;
        }
    }
    // ===================================================

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

    private static Object getCellValueAt(int rowNum, int colNum) {
        XSSFCell cell = getCellAt(rowNum, colNum);
        return getCellValueAt(cell);
    }

    private static Object getCellValueAt(XSSFCell cell) {
        return switch (cell.getCellType()) {
            case NUMERIC -> cell.getNumericCellValue();
            case STRING -> cell.getStringCellValue();
            case BLANK -> "";
            default -> cell.getRawValue();
        };
    }

    private static String getCellStringValueAt(int rowNum, int colNum) {
        return getCellValueAt(rowNum, colNum).toString();
    }

    private static String getCellStringValueAt(XSSFCell cell) {
        return getCellValueAt(cell).toString();
    }

    private static void saveFile() throws IOException {
        //Create file system using specific name
        FileOutputStream out = new FileOutputStream(FILE_NAME);

        //write operation workbook using file out object
        workbook.write(out);
        out.close();
    }
}
