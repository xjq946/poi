package com.example.poi.utils;

import com.example.poi.model.User;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.beans.BeanInfo;
import java.beans.IntrospectionException;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Method;
import java.net.URLEncoder;
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

    public static final String READ_METHOD= "readMethod";

    public static final String WRITE_METHOD= "writeMethod";


    /**
     * 读取excel表，封装数据在list集合中
     * @param path
     * @return
     * @throws IOException
     * @throws IllegalArgumentException
     */
    public static <T> List<T> readExcel(String path) throws IOException, IllegalArgumentException {
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

    /**
     * Read the Excel 2007版本及以上
     * @param path
     * @param <T>
     * @return
     * @throws IOException
     */
    public static <T> List<T> readXlsx(String path) throws IOException {
        try {
            InputStream is = new FileInputStream(path);
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
            T obj = null;
            List<T> list = new ArrayList<>();
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
                        obj = (T)Class.forName("com.example.poi.model.User").newInstance();
                        T result = setObject(obj, xssfRow, xssfSheet);
                        list.add(result);
                    }
                }
            }
            return list;
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * Read the Excel 2003
     * @param path the path of the Excel
     * @return
     * @throws IOException
     */
    public static <T> List<T> readXls(String path) throws IOException {
        try {
            InputStream is = new FileInputStream(path);
            HSSFWorkbook hssfWorkbook = new HSSFWorkbook(is);
            T obj = null;
            List<T> list = new ArrayList<>();
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
                        obj = (T)Class.forName("com.example.poi.model.User").newInstance();
                        T result = setObject(obj, hssfRow, hssfSheet);
                        list.add(result);
                    }
                }
            }
            return list;
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 获取excel表的文件类型，如xls、xlsx
     * @param path
     * @return
     */
    public static String getSuffix(String path) {
        if(StringUtils.isBlank(path)){
            return EMPTY;
        }
        int index = path.lastIndexOf(DOT);
        if (index == -1) {
            return EMPTY;
        }
        return path.substring(index + 1);
    }


    /**
     * 上传时为对象赋值
     * @param obj
     * @param row
     * @param sheet
     * @param <T>
     * @return
     */
    public static <T> T setObject(T obj,Row row,Sheet sheet){
        try {
            List<String> properties = getFirstRow(sheet);
            Map<String, Map<String, Method>> allMethodMap = getMethod(obj.getClass());
            Map<String, Method> writeMethodMap = allMethodMap.get(WRITE_METHOD);
            for(int colNum=0;colNum<properties.size();colNum++){
                Method method = writeMethodMap.get(properties.get(colNum));
                Class<?>[] parameterTypes = method.getParameterTypes();
                String simpleName = parameterTypes[0].getSimpleName();
                Cell cell = row.getCell(colNum);
                if(simpleName.equals("Date")){
                    SimpleDateFormat sdf=new SimpleDateFormat("yyyy/MM/dd");
                    String dateStr = cell.getStringCellValue();
                    Date date = sdf.parse(dateStr);
                    java.sql.Date sqlDate=new java.sql.Date(date.getTime());
                    method.invoke(obj,sqlDate);
                }else if(simpleName.equals("String")){
                    method.invoke(obj,cell.getStringCellValue());
                }else if(simpleName.equals("Boolean")){
                    method.invoke(obj,Boolean.parseBoolean(cell.getStringCellValue()));
                }else if(simpleName.equals("Integer")){
                    method.invoke(obj,Integer.parseInt(cell.getStringCellValue()));
                }else if(simpleName.equals("Long")){
                    method.invoke(obj,Long.parseLong(cell.getStringCellValue()));
                }else{
                    method.invoke(obj,Double.parseDouble(cell.getStringCellValue()));
                }
            }
            return obj;
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 下载时为单元格赋值
     * @param obj
     * @param row
     * @param sheet
     * @param <T>
     */
    public static <T> void setCellData(T obj,Row row,Sheet sheet){
        try {
            List<String> properties = getFirstRow(sheet);
            Map<String, Map<String, Method>> allMethodMap = getMethod(obj.getClass());
            Map<String, Method> readMethodMap = allMethodMap.get(READ_METHOD);
            for(int colNum=0;colNum<properties.size();colNum++){
                Method method = readMethodMap.get(properties.get(colNum));
                Class<?> returnType = method.getReturnType();
                String simpleName = returnType.getSimpleName();
                Cell cell = row.createCell(colNum);
                if(simpleName.equals("Date")){
                    java.sql.Date date = (java.sql.Date)method.invoke(obj);
                    SimpleDateFormat sdf=new SimpleDateFormat("yyyy-MM-dd");
                    String format = sdf.format(date);
                    cell.setCellValue(format);
                }else if(simpleName.equals("String")){
                    String str= (String)method.invoke(obj);
                    cell.setCellValue(str);
                }else if(simpleName.equals("Boolean")){
                    Boolean b= (Boolean)method.invoke(obj);
                    cell.setCellValue(b);
                }else if(simpleName.equals("Integer")){
                    Integer i = (Integer)method.invoke(obj);
                    cell.setCellValue(i);
                }else if(simpleName.equals("Long")){
                    Long aLong = (Long)method.invoke(obj);
                    cell.setCellValue(aLong);
                }else{
                    Double aDouble = (Double)method.invoke(obj);
                    cell.setCellValue(aDouble);
                }
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 将单元格的数据类型的值全部转换为字符串输出
     * @param cell
     * @return
     */
    private static String getValue(Cell cell) {
        if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
            return String.valueOf(cell.getBooleanCellValue());
        } else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
            return String.valueOf(cell.getNumericCellValue());
        } else {
            return String.valueOf(cell.getStringCellValue());
        }
    }

    /**
     * 下载，设置响应头
     * @param fileName
     * @param response
     * @param workbook
     */
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

    /**
     * 获取对象的所有属性，暂未使用
     * @param clazz
     * @param <T>
     * @return
     */
    public static <T> Map<String,String> getAllProperties(Class<T> clazz){
        try {
            BeanInfo beanInfo= Introspector.getBeanInfo(clazz);
            PropertyDescriptor[] propertyDescriptors = beanInfo.getPropertyDescriptors();
            Map<String,String> propertyMap=new LinkedHashMap<>();
            if(Objects.nonNull(propertyDescriptors)&& propertyDescriptors.length>0){
                for(PropertyDescriptor propertyDescriptor:propertyDescriptors){
                    String name = propertyDescriptor.getName();
                    propertyMap.put(name,name);
                }
            }
            return propertyMap;
        } catch (IntrospectionException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 利用内省获取类中的读写方法
     * @param clazz
     * @param <T>
     * @return
     */
    public static <T> Map<String,Map<String,Method>> getMethod(Class<T> clazz){
        try {
            BeanInfo beanInfo= Introspector.getBeanInfo(clazz);
            PropertyDescriptor[] propertyDescriptors = beanInfo.getPropertyDescriptors();
            Map<String,Map<String,Method>> allMethodMap=new LinkedHashMap<>();
            Map<String,Method> writeMethodMap=new LinkedHashMap<>();
            Map<String,Method> readMethodMap=new LinkedHashMap<>();
            if(Objects.nonNull(propertyDescriptors)&& propertyDescriptors.length>0){
                for(PropertyDescriptor propertyDescriptor:propertyDescriptors){
                    String name = propertyDescriptor.getName();
                    writeMethodMap.put(name,propertyDescriptor.getWriteMethod());
                    readMethodMap.put(name,propertyDescriptor.getReadMethod());
                }
            }
            allMethodMap.put(WRITE_METHOD,writeMethodMap);
            allMethodMap.put(READ_METHOD,readMethodMap);
            return allMethodMap;
        } catch (IntrospectionException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 获取excel表的第一行
     * @param sheet
     * @return
     */
    public static List<String> getFirstRow(Sheet sheet) {
        List<String> properties=new ArrayList<>();
        Row row = sheet.getRow(0);
        if(Objects.nonNull(row)){
            for(int colNum=0;colNum<row.getPhysicalNumberOfCells();colNum++){
                Cell cell = row.getCell(colNum);
                String value = getValue(cell);
                properties.add(value);
            }
        }
        return properties;
    }
}
