package JavaTool.ExcelMergeTool;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Objects;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * エクセルマージツール
 *
 */
public class App {
    private static String INPUT_DIR = "/Users/yoshiyuki.hayashi/Pictures/テーブル定義/";
    private static String OUTPUT_DIR = "/Users/yoshiyuki.hayashi/Pictures/テーブル定義/output/";

    private static int rowNum = 1;
    private static int cellNum = 0;

    public static void main(String[] args) {
        Workbook outputWorkbook = null;

        FileOutputStream out = null;
        try {
            // 一つのファイルに出力
            outputWorkbook = new HSSFWorkbook();

            Sheet outputSheet = outputWorkbook.createSheet();

            readFolder(new File(INPUT_DIR), outputSheet);

            // 出力するエクセルファイルを指定
            out = new FileOutputStream(OUTPUT_DIR + "table_work_2.xls");

            // 書き込み
            outputWorkbook.write(out);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                out.close();
                outputWorkbook.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * ディレクトリを再帰的に読む
     * @param folderPath
     * @param outputSheet
     * @throws IOException
     * @throws InvalidFormatException
     * @throws EncryptedDocumentException
     */
    private static void readFolder(File dir, Sheet outputSheet)
            throws EncryptedDocumentException, InvalidFormatException, IOException {
        File[] files = dir.listFiles();
        // ファイルリストをファイル名で照準に並べ替え
        java.util.Arrays.sort(files, new java.util.Comparator<File>() {
            public int compare(File file1, File file2) {
                return file1.getName().compareTo(file2.getName());
            }
        });
        if (files == null)
            return;
        for (File file : files) {
            if (!file.exists())
                continue;
            else if (file.isDirectory())
                System.out.println("SKIP:フォルダ");
            else if (file.isFile() && !".DS_Store".equals(file.getName()))
                execute(file, outputSheet);
        }
    }

    /**
     * ファイルごとに所定の処理を実行する
     * @param file
     * @param outputSheet
     * @throws IOException
     * @throws InvalidFormatException
     * @throws EncryptedDocumentException
     */
    private static void execute(File file, Sheet outputSheet)
            throws EncryptedDocumentException, InvalidFormatException, IOException {
        // 各ファイルの読み込み
        System.out.println(file.getName());
        Workbook workbook = WorkbookFactory.create(new File(INPUT_DIR + file.getName()));
        Sheet inputSheet = workbook.getSheetAt(0);
        String sheetName = inputSheet.getSheetName();

        Row outputRow = null;
        Cell outputCell = null;

        int inputRowNum = 7;
        Row row = inputSheet.getRow(inputRowNum);
        while (row != null && row.getCell(0).getCellType() != Cell.CELL_TYPE_BLANK) {
            int inputCellNum = 1;
            Cell cell = row.getCell(inputCellNum);

            cellNum = 0;

            outputRow = outputSheet.createRow(rowNum);
            outputCell = outputRow.createCell(cellNum);
            outputCell.setCellValue(sheetName);
            cellNum++;
            while (cell != null) {
                outputCell = outputRow.createCell(cellNum);
                getCellValue(cell, outputCell);
                cellNum++;
                inputCellNum++;
                cell = row.getCell(inputCellNum);
            }
            rowNum++;
            inputRowNum++;
            row = inputSheet.getRow(inputRowNum);
        }
    }

    /**
     * セルの値を型に合わせて出力対象セルへ設定する
     * @param cell
     * @param outputSheet
     * @throws IOException
     * @throws InvalidFormatException
     * @throws EncryptedDocumentException
     */
    private static void getCellValue(Cell cell, Cell outputCell) {
        Objects.requireNonNull(cell, "cell is null");

        int cellType = cell.getCellType();
        if (cellType == Cell.CELL_TYPE_BLANK) {
            //System.out.println("SKIP:BLANK");
        } else if (cellType == Cell.CELL_TYPE_BOOLEAN) {
            outputCell.setCellValue(cell.getBooleanCellValue());
        } else if (cellType == Cell.CELL_TYPE_ERROR) {
            throw new RuntimeException("Error cell is unsupported");
        } else if (cellType == Cell.CELL_TYPE_FORMULA) {
            outputCell.setCellValue(cell.getCellFormula());
        } else if (cellType == Cell.CELL_TYPE_NUMERIC) {
            if (DateUtil.isCellDateFormatted(cell)) {
                throw new RuntimeException("Date cell is unsupported");
            } else {
                outputCell.setCellValue(cell.getNumericCellValue());
            }
        } else if (cellType == Cell.CELL_TYPE_STRING) {
            outputCell.setCellValue(cell.getStringCellValue());
        } else {
            throw new RuntimeException("Unknow type cell");
        }
    }
}
