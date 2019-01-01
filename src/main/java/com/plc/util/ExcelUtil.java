package com.plc.util;

import com.alibaba.fastjson.JSON;
import com.plc.annotation.ExcelTitle;
import com.plc.config.ExcelConfig;
import org.apache.commons.beanutils.ConvertUtils;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.net.URL;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Created by Administrator on 2018/12/30.
 */
public class ExcelUtil {
    /**
     * 通过文件名读取文件内容
     * @param fileName
     * @return
     * @throws IOException
     */
    public static Map<String,Object> readFileByFileName(String fileName) throws IOException {
        ClassLoader classLoader = ExcelUtil.class.getClassLoader();
        /**
         getResource()方法会去classpath下找这个文件，获取到url resource, 得到这个资源后，调用url.getFile获取到 文件 的绝对路径
         */
        URL url = classLoader.getResource(fileName);
        if(null != url){
            File file = new File(url.getFile());
            String content= FileUtils.readFileToString(file,"UTF-8");
            Map maps = (Map) JSON.parse(content);
            return maps;
        }
        /**
         * url.getFile() 得到这个文件的绝对路径
         */
        //此处有问题，如果没有找到配置文件应读取默认配置参数
        return null;
    }

    public static Map getExcelConfig(String excelFileName) throws IOException{
        Map params = readFileByFileName(ExcelConfig.FILENAME);
        if(null == params){
            params = new HashMap();
            params.put("startRow",ExcelConfig.STARTROW);
            params.put("titleRow",ExcelConfig.TITLEROW);
            return params;
        }else{
            Map excelConfig = (Map) params.get(excelFileName);
            if(null == excelConfig){
                excelConfig = new HashMap();
                excelConfig.put("startRow",ExcelConfig.STARTROW);
                excelConfig.put("titleRow",ExcelConfig.TITLEROW);
            }
            return excelConfig;
        }
    }


    /**
     * 将读出的数据转成实体属性
     * @param rs
     * @param clazz
     * @return
     * @throws Exception
     */
    public static <T> T columnToField(Map<String,Object> rs, Class<T> clazz) throws Exception{
        //获取每行的数据的map  title->数据
        //根据title去实体类中寻找对应的字段，并把赋值
        //先找到对应的字段
        //通过反射将数据赋值
        if(rs!=null && rs.size() > 0){
            T t = null;
            for (String str : rs.keySet()) {
                Field field = getField(str,clazz);
                if(null != field){
                    if(t == null){
                        t = clazz.newInstance();
                    }
                    String name = field.getName();
                    Class type = field.getType();
                    Method method = null;
                    try {
                        //反射得到javaBean每个字段的set方法
                        method = clazz.getMethod("set" + name.replaceFirst(name.substring(0, 1),
                                name.substring(0, 1).toUpperCase()), type);
                        method.invoke(t, ConvertUtils.convert(rs.get(str), type));
                    } catch (Exception e) {
                        e.printStackTrace();
                        throw new Exception("Import excel read error!");
                    }
                }
            }
            return t;
        }
        return null;
    }


    /**
     * 封装单元格的值
     * @param titleMap
     * @param index
     * @param cellValue
     * @return
     */
    public static Map<String, Object> setStandardData(Map<Integer, Object> titleMap,
                                                Integer index, String cellValue) {
        Map<String, Object> temp = new HashMap<>();
        String field = (String) titleMap.get(index);
        if (field != null && !field.isEmpty()) {
            temp.put(field, cellValue);
            return temp;
        } else {
            return null;
        }
    }

    /**
     * 合并单元格处理,获取合并行
     * @param sheet
     * @return
     */
    public static List<CellRangeAddress> getCombineCell(Sheet sheet) {
        List<CellRangeAddress> list = new ArrayList<CellRangeAddress>();
        // 获得一个 sheet 中合并单元格的数量
        int sheetmergerCount = sheet.getNumMergedRegions();
        // 遍历所有的合并单元格
        for (int i = 0; i < sheetmergerCount; i++) {
            // 获得合并单元格保存进list中
            CellRangeAddress ca = sheet.getMergedRegion(i);
            list.add(ca);
        }
        return list;
    }



    /**
     * 获取单元格的值
     *
     * @param cell
     * @return
     */
    public static String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        String cellValue = "";
        int cellType = cell.getCellType();
        switch (cellType) {
            case Cell.CELL_TYPE_STRING: // 文本
                cellValue = cell.getStringCellValue();
                break;
            case Cell.CELL_TYPE_NUMERIC: // 数字、日期
                if (DateUtil.isCellDateFormatted(cell)) {
                    System.out.println(cell.getDateCellValue());
                    System.out.println(cell.getCellStyle().getDataFormat());
                    System.out.println(cell.getCellStyle().getDataFormatString());

                    SimpleDateFormat sdf = null;
                    if (cell.getCellStyle().getDataFormat() == HSSFDataFormat.getBuiltinFormat("h:mm")) {
                        sdf = new SimpleDateFormat("HH:mm");
                    } else if (cell.getCellStyle().getDataFormat() == HSSFDataFormat
                            .getBuiltinFormat("yyyy/MM/dd HH:mm:ss")) {
                        sdf = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
                    }else if (cell.getCellStyle().getDataFormat() == 177
                            || cell.getCellStyle().getDataFormat() == 176
                            || cell.getCellStyle().getDataFormat() == 22
                            || cell.getCellStyle().getDataFormat() == 58) {
                        sdf = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
                    }else {// 日期
                        sdf = new SimpleDateFormat("yyyy/MM/dd");
                    }
                    Date date = cell.getDateCellValue();
                    cellValue = sdf.format(date);
                } else {
                    // 返回数值类型的值
                    Object inputValue = null;// 单元格值
                    Long longVal = Math.round(cell.getNumericCellValue());
                    Double doubleVal = cell.getNumericCellValue();
                    if(Double.parseDouble(longVal + ".0") == doubleVal){   //判断是否含有小数位.0
                        inputValue = longVal;
                    }
                    else{
                        inputValue = doubleVal;
                    }
                    DecimalFormat df = new DecimalFormat("#.####");    //格式化为四位小数，按自己需求选择；
                    cellValue = String.valueOf(df.format(inputValue));      //返回String类型
                }
                break;
            case Cell.CELL_TYPE_BOOLEAN: // 布尔型
                cellValue = String.valueOf(cell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_BLANK: // 空白
                cellValue = cell.getStringCellValue();
                break;
            case Cell.CELL_TYPE_ERROR: // 错误
                cellValue = "";
                break;
            case Cell.CELL_TYPE_FORMULA: // 公式
                cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;
            default:
                cellValue = "";
        }
        return cellValue;
    }

    /**
     * 判断单元格是否为合并单元格，是的话则将单元格的值返回
     *
     * @param listCombineCell
     *            存放合并单元格的list
     * @param cell
     *            需要判断的单元格
     * @param sheet
     *            sheet
     * @return
     */
    public static String isCombineCell(List<CellRangeAddress> listCombineCell, Cell cell, Sheet sheet)
            throws Exception {
        int firstC = 0;
        int lastC = 0;
        int firstR = 0;
        int lastR = 0;
        String cellValue = null;
        for (CellRangeAddress ca : listCombineCell) {
            // 获得合并单元格的起始行, 结束行, 起始列, 结束列
            firstC = ca.getFirstColumn();
            lastC = ca.getLastColumn();
            firstR = ca.getFirstRow();
            lastR = ca.getLastRow();
            if (cell.getRowIndex() >= firstR && cell.getRowIndex() <= lastR) {
                if (cell.getColumnIndex() >= firstC && cell.getColumnIndex() <= lastC) {
                    Row fRow = sheet.getRow(firstR);
                    Cell fCell = fRow.getCell(firstC);
                    cellValue = getCellValue(fCell);
                    break;
                }
            } else {
                cellValue = "";
            }
        }
        return cellValue;
    }

    /**
     * 获取实体中对应的字段
     * @param value
     * @param c
     * @return
     */
    public static Field getField(String value, Class c) {
        Field[] declaredFields = c.getDeclaredFields();
        Field field = null;
        for (Field f : declaredFields) {
            ExcelTitle excelTitle = f.getAnnotation(ExcelTitle.class);
            if (excelTitle != null) {
                String[] columnName = excelTitle.value();
                for (int i = 0; i < columnName.length; i++){
                    if (columnName[i].equals(value)) {
                        field = f;
                        break;
                    }
                }
            }
        }
        return field;
    }

    /**
     * 判断指定的单元格是否是合并单元格
     * @param sheet
     * @param row
     * @param column
     * @return
     */
    public static boolean isMergedRegion(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    return true;
                }
            }
        }
        return false;
    }
}
