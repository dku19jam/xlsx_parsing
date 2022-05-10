import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Date;
import java.util.HashMap;
import java.text.SimpleDateFormat;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static java.lang.System.out;
import static org.junit.Assert.*;


public class parsing {

    private String fileName;
    private final String ExcelPath = "C:\\Program Files (x86)\\KEPRI\\통합운영프로그램\\Excel\\24260157644";

    private FileInputStream file;

    public parsing() throws IOException{
        String filepath = setPath();
        out.println("path : " + filepath);

        this.file = new FileInputStream(filepath);
    }

    private String setPath() {
        File directory = new File(ExcelPath);
        File[] files = directory.listFiles(File::isFile);
        long lastModifiedTime = Long.MIN_VALUE;
        File chosenFile = null;

        if (files != null){
            for (File file: files) {
                if(file.lastModified() > lastModifiedTime){
                    chosenFile = file;
                    lastModifiedTime = file.lastModified();
                }
            }
        }
        if (chosenFile != null){
            return chosenFile.getPath();
        }
        throw new Error("xlsx file what you were looking for does not exist");
    }


    public parsing(String fileName) {
        this.fileName = fileName;
    }

    public void getCellDataByColumnName(String columName) throws IOException {
        XSSFWorkbook workbook = null;
        workbook = new XSSFWorkbook(file);
        String value = "";
        XSSFSheet sheet = workbook.getSheetAt(0);
        int rows = sheet.getPhysicalNumberOfRows(); // 행의 수

        for (int r = 0; r < rows; r++) {
            XSSFRow row = sheet.getRow(r); // 0 ~ rows
            XSSFCell cell = row.getCell(0);
            if (cell.getStringCellValue().equals(columName)) {
                cell = row.getCell(1);
                switch (cell.getCellType()) {
                    case FORMULA -> value = cell.getCellFormula();
                    case NUMERIC -> value = cell.getNumericCellValue() + "";
                    case STRING -> value = cell.getStringCellValue() + "";
                    case BLANK -> value = cell.getBooleanCellValue() + "";
                    case ERROR -> value = cell.getErrorCellValue() + "";
                }
                out.println(columName + "값은 " + value);
            }
        }
    }

    public void getDataForGivenDays(int day) throws IOException {
        HashMap<String, Integer> type = new HashMap<String, Integer>();
        type.put("평균전압/전류",10);
        type.put("LP데이터",96);
        int cnt = 0;
        XSSFWorkbook workbook = null;
        workbook = new XSSFWorkbook(file);

        int cells = 0;
        XSSFSheet sheet = workbook.getSheetAt(0);
        while (true){
            XSSFRow row = sheet.getRow(cnt++);
            XSSFCell cell = row.getCell(0);
            if (cell.getStringCellValue().equals("일자/시간")) {
                for (int i = 0; i < row.getPhysicalNumberOfCells(); i++) {
                    XSSFCell category = row.getCell(i);
                    if (category.getCellType() == CellType.BLANK) {
                        cells = i;
                        break;
                    }
                }
                break;
            }
        }
        for (int r = cnt; r < cnt+ day * type.get("평균전압/전류"); r++){
            XSSFRow row = sheet.getRow(r); // 0 ~ rows
            for (int c = 0 ; c < cells ; c++){
                XSSFCell cell = row.getCell(c);
                String value = "";
                if (cell == null) { // r열 c행의 cell이 비어있을 때
                    continue;
                } else { // 타입별로 내용 읽기
                    switch (cell.getCellType()) {
                        case FORMULA -> value = cell.getCellFormula();
                        case NUMERIC -> value = cell.getNumericCellValue() + "";
                        case STRING -> value = cell.getStringCellValue() + "";
                        case BLANK -> value = cell.getBooleanCellValue() + "";
                        case ERROR -> value = cell.getErrorCellValue() + "";
                    }
                }
                if (c == 0){
                    Date date = HSSFDateUtil.getJavaDate(cell.getNumericCellValue());
                    long time = date.getTime();
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                    value = sdf.format(time);

                }
                System.out.println(r + "번 행 : " + c + "번 열 값은: " + value);
                out.println(cell.getCellType());
            }
        }
    }


//        FileInputStream file = new FileInputStream("C:\\Users\\TB-NTB-118\\Desktop\\08250143205_LP정보_220426024503.xlsx"); // 파일 읽기
//        XSSFWorkbook workbook = new XSSFWorkbook(file); // 엑셀 파일 파싱
//
//        XSSFSheet sheet = workbook.getSheetAt(0); // 엑셀 파일의 첫번째 (0) 시트지
//        int rows = sheet.getPhysicalNumberOfRows(); // 행의 수
//
//        for (int r = 1; r < 97; r++) {
//            XSSFRow row = sheet.getRow(r); // 0 ~ rows
//
//            for (int c = 0; c < 10; c++) {
//                XSSFCell cell = row.getCell(c); // 0 ~ cell
//                String value = "";
//
//                if (cell == null) { // r열 c행의 cell이 비어있을 때
//                    continue;
//                } else { // 타입별로 내용 읽기
//                    switch (cell.getCellType()) {
//                        case FORMULA:
//                            value = cell.getCellFormula();
//                            break;
//                        case NUMERIC:
//                            value = cell.getNumericCellValue() + "";
//                            break;
//                        case STRING:
//                            value = cell.getStringCellValue() + "";
//                            break;
//                        case BLANK:
//                            value = cell.getBooleanCellValue() + "";
//                            break;
//                        case ERROR:
//                            value = cell.getErrorCellValue() + "";
//                            break;
//                    }
//
//                }
//
//                System.out.println(r + "번 행 : " + c + "번 열 값은: " + value);
//
//            }
//
//        }
//    }

}
