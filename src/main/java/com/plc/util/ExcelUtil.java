package com.plc.util;

import com.plc.annotation.ExcelField;
import com.plc.config.ExcelConfig;
import com.plc.model.ExcelCustomData;
import com.plc.model.MergeColumn;
import org.apache.commons.beanutils.ConvertUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;

import java.beans.PropertyDescriptor;
import java.lang.reflect.*;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;

/**
 * Created by Administrator on 2018/12/30.
 */
public class ExcelUtil {

    /**
     * 将读出的数据转成实体属性
     * @param rs
     * @param clazz
     * @return
     * @throws Exception
     */
    public <T> T columnToField(Map<String,Object> rs, Class<T> clazz) throws Exception{
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
    public Map<String, Object> setStandardData(Map<Integer, Object> titleMap,
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
    public List<CellRangeAddress> getCombineCell(Sheet sheet) {
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
    public String getCellValue(Cell cell) {
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
    public String isCombineCell(List<CellRangeAddress> listCombineCell, Cell cell, Sheet sheet)
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
    public Field getField(String value, Class c) {
        Field[] declaredFields = c.getDeclaredFields();
        Field field = null;
        for (Field f : declaredFields) {
            ExcelField excelField = f.getAnnotation(ExcelField.class);
            if (excelField != null) {
                String[] columnName = excelField.columnName();
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
    public boolean isMergedRegion(Sheet sheet, int row, int column) {
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

    //-------------------------------导出---------------------------------

    public <T> void writeExcel(XSSFWorkbook wb, Sheet sheet, List<T> data , Class<T> clazz,ExcelConfig excelConfig,ExcelStyle excelStyle) throws Exception {

        int rowIndex = excelConfig.getWriteStartRow();
        //读取excel配置文件内自定义和合并单元格
        writeCustomExcelData(wb,sheet,excelConfig,excelStyle);

        //获取排序后的字段
        List<Field> fieldList = getFieldListBySort(clazz);
        rowIndex = writeTitlesToExcel(wb, sheet, fieldList,excelConfig,excelStyle);
        writeRowsToExcel(wb, sheet, data,fieldList,clazz,excelStyle,rowIndex);
        autoSizeColumns(sheet, data.size() + 1);

    }

    public <T> void writeExcel(XSSFWorkbook wb, Sheet sheet, List<T> data , Class<T> clazz,ExcelConfig excelConfig) throws Exception {

        int rowIndex = excelConfig.getWriteStartRow();
        //获取排序后的字段
        List<Field> fieldList = getFieldListBySort(clazz);
        //获取Excel样式
        ExcelStyle excelStyle = new ExcelStyle();
        rowIndex = writeTitlesToExcel(wb, sheet, fieldList,excelConfig,excelStyle);
        writeRowsToExcel(wb, sheet, data,fieldList,clazz,excelStyle,rowIndex);
        autoSizeColumns(sheet, data.size() + 1);

    }

    //获取排序后的字段
    public <T> List<Field> getFieldListBySort(Class<T> classz) throws Exception {
        List<Field> fieldList = new ArrayList<>();
        getAnnotationFields(fieldList, classz);
        for (int i = 0; i < fieldList.size(); i++) {
            for (int j = i + 1; j < fieldList.size(); j++) {
                ExcelField column_i = fieldList.get(i).getAnnotation(ExcelField.class); //获取指定类型注解
                ExcelField column_j = fieldList.get(j).getAnnotation(ExcelField.class); //获取指定类型注解
                if (column_i != null && column_j != null) {
                    if (column_i.order() > column_j.order()) {
                        Field temp = fieldList.get(i);
                        fieldList.set(i, fieldList.get(j));
                        fieldList.set(j, temp);
                    }
                }
            }
        }
        return fieldList;
    }

    /**
     * 写标题
     *
     * @param wb
     * @param sheet
     * @param fieldList
     * @return
     */
    private int writeTitlesToExcel(XSSFWorkbook wb, Sheet sheet, List<Field> fieldList,ExcelConfig excelConfig,ExcelStyle excelStyle) {
        int rowIndex = excelConfig.getWriteStartRow();
        int colIndex = excelConfig.getWriteStartCol();

        //获取表头的样式
        XSSFCellStyle titleStyle = excelStyle.setTitleStyle(wb);

        Row titleRow = sheet.createRow(rowIndex);
        // titleRow.setHeightInPoints(25);
        colIndex = 0;

        for (Field field : fieldList) {
            ExcelField excelAnnotation = field.getAnnotation(ExcelField.class); //获取指定类型注解
            String[] columnNameArr = excelAnnotation.columnName();
            for (int i = 0; i < columnNameArr.length; i++) {
                Cell cell = titleRow.createCell(colIndex);
                if (excelAnnotation.columnWidth() != 0) {
                    sheet.setColumnWidth(colIndex, excelAnnotation.columnWidth());
                }
                cell.setCellValue(columnNameArr[i]);
                cell.setCellStyle(titleStyle);
                colIndex++;
            }
        }

        rowIndex++;
        return rowIndex;
    }

    /**
     * 写内容
     *
     * @param wb
     * @param sheet
     * @param data      数据列表
     * @param fieldList
     * @param rowIndex
     * @param <T>
     * @return
     */
    private <T> int writeRowsToExcel(XSSFWorkbook wb, Sheet sheet, List<T> data, List<Field> fieldList,
                                            Class<T> clazz,ExcelStyle excelStyle, int
                                                    rowIndex) throws Exception {
        int colIndex = 0;

        //设置数据的样式
        XSSFCellStyle dataStyle = excelStyle.setDataStyle(wb);


        //存放合并单元格信息
//        Map<String,Integer> mergeColumn = new HashMap<>();
        for (T entity : data) {
            //结果集合
            List<Map<String, Object>> resultList = new ArrayList<>();
            //存放每个字段和对应的值
            Map<String, Object> paramMap = new HashMap<>();
            //递归解析数据，将解析完的数据存入resultList
            analysisData(paramMap, resultList, fieldList, entity, clazz, 0);
            //获取单元格合并信息
            List<MergeColumn> mergeColumnList = getMergeColumn(resultList);
            //合并单元格
            //创建map记录每列Merge的行数
            Map<String, Integer> mergeRowIndex = new HashMap<>();
            //将解析的数据写入excel
            for (Map<String, Object> map : resultList) {

                Row dataRow = sheet.createRow(rowIndex);
                // dataRow.setHeightInPoints(25);
                colIndex = 0;
                for (Field field : fieldList) {
                    ExcelField excelAnnotation = field.getAnnotation(ExcelField.class); //获取指定类型注解
                    String[] columnNameArr = excelAnnotation.columnName();
                    for (int i = 0; i < columnNameArr.length; i++) {
                        Object entityValue = null;
                        entityValue = map.get(field.getDeclaringClass() + "." + field.getName());
                        Cell cell = dataRow.createCell(colIndex);
                        String value;

                        if (excelAnnotation.isRowMerger()) {
                            if (mergeColumnList != null) {
                                for (MergeColumn mergeColumn : mergeColumnList) {
                                    if ((field.getDeclaringClass() + "." + field.getName()).equals(mergeColumn.getColumn())) {
                                        if (mergeRowIndex.get(mergeColumn.getColumn()) == null) {
                                            mergeRowIndex.put(mergeColumn.getColumn(), rowIndex);
                                        }
                                        if (rowIndex < mergeRowIndex.get(mergeColumn.getColumn())) {
                                            continue;
                                        }
                                        //MergeCount的数量包括当前行，与html rowspan使用方式一样
                                        sheet.addMergedRegion(new CellRangeAddress(mergeRowIndex.get(mergeColumn.getColumn()), mergeRowIndex.get(mergeColumn.getColumn()) + mergeColumn.getMergeCount() - 1,
                                                colIndex, colIndex));
                                        mergeRowIndex.put(mergeColumn.getColumn(), mergeRowIndex.get(mergeColumn.getColumn()) + mergeColumn.getMergeCount());
                                        mergeColumnList.remove(mergeColumn);
                                        break;
                                    }
                                }
                            }
                        }
                        if (excelAnnotation.isColMerger()) {
//                            if (mergeColumnList != null) {
//                                for (MergeColumn mergeColumn : mergeColumnList) {
//                                    if ((field.getDeclaringClass() + "." + field.getName()).equals(mergeColumn.getColumn())) {
//                                        //MergeCount的数量包括当前行，与html rowspan使用方式一样
//                                        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, colIndex, colIndex + mergeColumn
//                                                .getMergeCount() - 1));
//                                    }
//                                }
//                            }
                        }
                        if (excelAnnotation.isDate()) {
                            if (entityValue != null && Long.parseLong(entityValue.toString()) != 0) {
                                String fomate = excelAnnotation.fomat();
                                SimpleDateFormat f = new SimpleDateFormat(fomate);
                                value = f.format(new Date(Long.parseLong(entityValue.toString())));
                                cell.setCellValue(value);
                            }
                        } else if (excelAnnotation.isNum()
                                && (entityValue != null && !"".equals(entityValue.toString().trim()))) {
                            cell.setCellValue(new BigDecimal(entityValue.toString()).setScale(2, BigDecimal.ROUND_HALF_UP).toString());
                        } else if (excelAnnotation.isInteger()
                                && (entityValue != null && !"".equals(entityValue.toString().trim()))) {
                            cell.setCellValue(Integer.parseInt(entityValue.toString()));
                        } else {
                            if (entityValue != null) {
                                cell.setCellValue(entityValue.toString());
                            } else {
                                cell.setCellValue("");
                            }
                        }
                        cell.setCellStyle(dataStyle);
                        colIndex++;
                    }
                }
                rowIndex++;
            }
        }
        return rowIndex;
    }

    /**
     * 获取可以合并的·单元格
     * @param resultList
     * @param <T>
     * @return
     */
    private <T> List<MergeColumn> getMergeColumn(List<Map<String, Object>> resultList) {
        List<MergeColumn> mergeColumnList = new ArrayList<>();
        //临时变量，存放每个字段出现的次数
        Map<String, Integer> m = new HashMap<>();
        //上一条数据
        Map<String, Object> lastData = new HashMap<>();
        for (Map<String, Object> map : resultList) {
            for (String key : map.keySet()) {
                if (m.get(key + "=" + map.get(key)) == null) {
                    m.put(key + "=" + map.get(key), 1);
                } else {
                    if (map.get(key).equals(lastData.get(key))) {
                        m.put(key + "=" + map.get(key), m.get(key + "=" + map.get(key).toString()) + 1);
                    }
                }
            }
            lastData = map;
        }

        for (String key : m.keySet()) {
            if (m.get(key) > 1) {
                String[] arr = key.split("=");
                MergeColumn mergeColumn = new MergeColumn();
                mergeColumn.setColumn(arr[0]);
                mergeColumn.setMergeCount(m.get(key));
                mergeColumnList.add(mergeColumn);
            }
        }
        return mergeColumnList;
    }

    private <T> void analysisData(Map<String, Object> paramMap, List<Map<String, Object>> resultList, List<Field> fieldList, T t, Class
            clazz, int currentLevel) throws Exception {
        //总层数
        int totalLevel = 0;
        for (Field field : fieldList) {
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            if (excelField.isObject()) {
                totalLevel++;
            }
        }
        Field[] fields = clazz.getDeclaredFields();
        AccessibleObject.setAccessible(fields, true);
        for (int j = 0; j < fieldList.size(); j++) {
            for (Field f : fields) {
                if ((fieldList.get(j).getDeclaringClass() + fieldList.get(j).getName()).equals(f.getDeclaringClass() + f.getName())) {
                    ExcelField excelField = fieldList.get(j).getAnnotation(ExcelField.class);
                    if (excelField != null) {
                        //解析数据
                        Object object = null;
                        PropertyDescriptor pd = new PropertyDescriptor(f.getName(), clazz);
                        Method getMethod = pd.getReadMethod();
                        if (getMethod.invoke(t) != null) {
                            object = getMethod.invoke(t);
                        }
                        if (!excelField.isObject()) {
                            paramMap.put(t.getClass() + "." + f.getName(), object);
                        }
                        if (excelField.isObject()) {
                            //当前层数
                            currentLevel++;
                            if (object instanceof List) {
                                for (Object obj : (List) object) {
                                    analysisData(paramMap, resultList, fieldList, obj, obj.getClass(), currentLevel);
                                    if (currentLevel == totalLevel) {
                                        Map<String, Object> m = new HashMap<>();
                                        for (String key : paramMap.keySet()) {
                                            m.put(key, paramMap.get(key));
                                        }
                                        resultList.add(m);
                                    }
                                }
                            } else {
                                if (null != object) {
                                    analysisData(paramMap, resultList, fieldList, object, object.getClass(), currentLevel);
                                }
                                if (currentLevel == totalLevel) {
                                    Map<String, Object> m = new HashMap<>();
                                    for (String key : paramMap.keySet()) {
                                        m.put(key, paramMap.get(key));
                                    }
                                    resultList.add(m);
                                }

                            }
                        }
                        //如果就一层，直接添加到集合中
                        if (totalLevel == 0) {
                            if (j == fieldList.size() - 1) {
                                resultList.add(paramMap);
                            }
                        }
                    }
                }
            }
        }
    }

    private void getAnnotationFields(List<Field> fieldList, Class classz) throws Exception {
        Field[] fields = classz.getDeclaredFields();
        AccessibleObject.setAccessible(fields, true);
        for (Field field : fields) {
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            if (excelField != null) {
                fieldList.add(field);
                if (!isBaseType(field)) {
                    if (excelField.isObject()) {
                        if (field.getType() == List.class) {
                            // 如果是List类型，得到其Generic的类型
                            Type genericType = field.getGenericType();
                            if (genericType == null) continue;
                            // 如果是泛型参数的类型
                            if (genericType instanceof ParameterizedType) {
                                ParameterizedType pt = (ParameterizedType) genericType;
                                //得到泛型里的class类型对象
                                Class<?> clazz = (Class<?>) pt.getActualTypeArguments()[0];
                                getAnnotationFields(fieldList, clazz);
                            }
                        } else {
                            Type type = field.getType();
                            Class<?> clazz = Class.forName(((Class) type).getName());
                            getAnnotationFields(fieldList, clazz);
                        }
                    }
                }
            }
        }
    }

    private boolean isBaseType(Field f) {
        String[] types = {"java.lang.Integer",
                "java.lang.Double",
                "java.lang.Float",
                "java.lang.Long",
                "java.lang.Short",
                "java.lang.Byte",
                "java.lang.Boolean",
                "java.lang.Character",
                "java.lang.String",
                "int", "double", "long", "short", "byte", "boolean", "char", "float"};
        for (String str : types) {
            if (f.getType().getName().equals(str)) {
                return true;
            }
        }
        return false;
    }

    private void autoSizeColumns(Sheet sheet, int columnNumber) {

        for (int i = 0; i < columnNumber; i++) {
            int orgWidth = sheet.getColumnWidth(i);
            sheet.autoSizeColumn(i, true);
            int newWidth = (int) (sheet.getColumnWidth(i) + 100);
            if (newWidth > orgWidth) {
                sheet.setColumnWidth(i, newWidth);
            } else {
                sheet.setColumnWidth(i, orgWidth);
            }
        }
    }

    private void writeCustomExcelData(XSSFWorkbook wb,Sheet sheet,ExcelConfig excelConfig,ExcelStyle excelStyle){
        List<ExcelCustomData> excelCustomDatas = excelConfig.getExcelCustomDatas();
        if(null != excelCustomDatas){
            for (ExcelCustomData excelCustomData : excelCustomDatas) {
                CellRangeAddress cellRangeAddress = excelCustomData.getCellRangeAddress();
                sheet.addMergedRegion(cellRangeAddress);
                Row dataRow = sheet.createRow(cellRangeAddress.getFirstRow());
                Cell cell = dataRow.createCell(cellRangeAddress.getFirstColumn());
                cell.setCellValue(excelCustomData.getExcelText());
                XSSFCellStyle cellStyle = excelStyle.getDetailStyle(wb);
                cell.setCellStyle(cellStyle);
                // 使用RegionUtil类为合并后的单元格添加边框
                RegionUtil.setBorderBottom(1, cellRangeAddress, sheet); // 下边框
                RegionUtil.setBorderLeft(1, cellRangeAddress, sheet); // 左边框
                RegionUtil.setBorderRight(1, cellRangeAddress, sheet); // 有边框
                RegionUtil.setBorderTop(1, cellRangeAddress, sheet); // 上边框
            }
        }
    }

}
