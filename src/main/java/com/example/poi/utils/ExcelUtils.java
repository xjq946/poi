package com.example.poi.utils;

import com.example.poi.model.User;
import org.apache.commons.lang.StringUtils;
import org.apache.commons.lang.time.DateUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.format.CellFormatType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URLEncoder;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

public class ExcelUtils {

    /**
     * suffix of excel 2003
     */
    public static final String OFFICE_EXCEL_V2003_SUFFIX = "xls";
    /**
     * suffix of excel 2007
     */
    public static final String OFFICE_EXCEL_V2007_SUFFIX = "xlsx";

    public static final String NOT_EXCEL_FILE = " is Not a Excel file!";

    public static final String EMPTY = "";

    public static final String DOT = ".";


    public static List<User> readExcel(String path) throws IOException, IllegalArgumentException {
        if (StringUtils.isBlank(path)) {
            throw new IllegalArgumentException(path + " excel file path is either null or empty");
        } else {
            String suffix = getSuffix(path);
            if(StringUtils.isBlank(suffix)){
                throw new IllegalArgumentException(path + " suffix is either null or empty");
            }
            if (OFFICE_EXCEL_V2003_SUFFIX.equals(suffix)) {
                return readXls(path);
            } else if (OFFICE_EXCEL_V2007_SUFFIX.equals(suffix)) {
                return readXlsx(path);
            } else {
                throw new IllegalArgumentException(path + NOT_EXCEL_FILE);
            }
        }
    }

    public static List<User> readXlsx(String path) throws IOException {
        InputStream is = new FileInputStream(path);
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
        User user = null;
        List<User> list = new ArrayList<User>();
        // Read the Sheet
        for (int numSheet = 0; numSheet < xssfWorkbook.getNumberOfSheets(); numSheet++) {
            XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(numSheet);
            if (xssfSheet == null) {
                continue;
            }
            // Read the Row
            for (int rowNum = 1; rowNum <= xssfSheet.getLastRowNum(); rowNum++) {
                XSSFRow xssfRow = xssfSheet.getRow(rowNum);
                if (xssfRow != null) {
                    user=new User();
                    setUser(user,xssfRow);
                    list.add(user);
                }
            }
        }
        return list;
    }

    /**
     * Read the Excel 2003
     * @param path the path of the Excel
     * @return
     * @throws IOException
     */
    public static  List<User> readXls(String path) throws IOException {
        InputStream is = new FileInputStream(path);
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook(is);
        User user = null;
        List<User> list = new ArrayList<User>();
        // Read the Sheet
        for (int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++) {
            HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
            if (hssfSheet == null) {
                continue;
            }
            // Read the Row
            for (int rowNum = 1; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
                HSSFRow hssfRow = hssfSheet.getRow(rowNum);
                if (hssfRow != null) {
                    user=new User();
                    setUser(user,hssfRow);
                    list.add(user);
                }
            }
        }
        return list;
    }

    public static String getSuffix(String path) {
        if(StringUtils.isBlank(path)){
            return EMPTY;
        }
        int index = path.lastIndexOf(DOT);
        if (index == -1) {
            return EMPTY;
        }
        return path.substring(index + 1, path.length());
    }

    public static void setUser(User user, Row row){
        user.setId((int)Float.parseFloat(getValue(row.getCell(0))));
        user.setName(row.getCell(1).getStringCellValue());
        user.setAge(getValue(row.getCell(2)).substring(0,getValue(row.getCell(2)).lastIndexOf(".")));

        try {
            SimpleDateFormat sdf=new SimpleDateFormat("yyyy/MM/dd");
            Date date = row.getCell(3).getDateCellValue();
            user.setBirthday(new java.sql.Date(date.getTime()));
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    private static String getValue(Cell cell) {
        if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
            return String.valueOf(cell.getBooleanCellValue());
        } else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
            return String.valueOf(cell.getNumericCellValue());
        } else {
            return String.valueOf(cell.getStringCellValue());
        }
    }

    public static void downLoadExcel(String fileName, HttpServletResponse response, Workbook workbook) {
        try {
            response.setCharacterEncoding("UTF-8");
            response.setHeader("content-Type", "application/vnd.ms-excel");
            response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
            workbook.write(response.getOutputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
