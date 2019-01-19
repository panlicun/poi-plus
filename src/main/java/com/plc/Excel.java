package com.plc;

import com.plc.config.ExcelConfig;
import com.plc.constant.Constant;
import com.plc.util.ExcelUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.util.*;

public class Excel<T> {

    public static <T> List<T> readExcelToList(String excelFileName, InputStream inputStream, Class<T> clazz) throws Exception {
        //读取classpath下的json文件
        ExcelConfig config = ExcelUtil.getExcelConfig(excelFileName);
        int startRowNo = config.getReadStartRow();
        int titleRowNo = config.getReadTitleRow();
        //存放表头列
        Map<Integer, Object> titleMap = new HashMap<>();
        //返回的数据集合
        List<T> lists = new ArrayList<>();
        Workbook workbook = WorkbookFactory.create(inputStream); // 这种方式 Excel 2003/2007/2010 都是可以处理的
        int sheetCount = workbook.getNumberOfSheets(); // Sheet的数量
        for (int s = 0; s < sheetCount; s++) {   //遍历每一个sheet
            Sheet sheet = workbook.getSheetAt(s);
            if (null != sheet) {
                int rowCount = sheet.getPhysicalNumberOfRows(); // 获取总行数
                // 合并单元格处理,获取合并行
                List<CellRangeAddress> cras = ExcelUtil.getCombineCell(sheet);
                for (int r = 0; r < rowCount; r++) {
                    if (r < startRowNo && r != titleRowNo) {
                        continue;
                    }
                    Row row = sheet.getRow(r);
                    int cellCount = row.getPhysicalNumberOfCells();  //获取总列数
                    //用来存放每行的数据
                    Map<String, Object> temp = new HashMap<>();
                    for (int c = 0; c < cellCount; c++) {
                        Cell cell = row.getCell(c);
                        if (null == cell) {
                            cellCount++;
                            continue;
                        }
                        String cellValue = ExcelUtil.getCellValue(cell);
                        //判断当前cell是否是合并单元格
                        if (ExcelUtil.isMergedRegion(sheet, r, c)) {
                            cellValue = ExcelUtil.isCombineCell(cras, cell, sheet);
                        }
                        if (r == titleRowNo) {  //读取标题
                            titleMap.put(c, cellValue);
                        } else {  //处理数据
                            Map<String, Object> map = ExcelUtil.setStandardData(titleMap, c,
                                    cellValue);
                            if (map != null) {
                                temp.putAll(map);
                            }
                        }
                    }
                    if (r >= startRowNo) {
                        //转换成实体
                        T t = ExcelUtil.columnToField(temp, clazz);
                        if(null != t){
                            lists.add(t);
                        }
                    }
                }
            }
        }
        workbook.close();
        return lists;
    }

    public static <T> void exportExcel(HttpServletResponse response, String fileName, Map<String, Map<Class, List>> data) throws Exception {
        // 告诉浏览器用什么软件可以打开此文件
        response.setHeader("content-Type", "application/vnd.ms-excel");
        // 下载文件的默认名称
        response.setHeader("Content-Disposition", "attachment;filename="+ URLEncoder.encode(fileName, "utf-8"));
        exportExcel(data, response.getOutputStream());
    }

    public static <T> void exportExcel(Map<String, Map<Class, List>> data, OutputStream out) throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        try {
            for (String key : data.keySet()) {
                String sheetName = key;
                XSSFSheet sheet = wb.createSheet(sheetName);
                for (Class clazz : data.get(key).keySet()) {
                    ExcelUtil.writeExcel(wb, sheet, data.get(key).get(clazz),clazz);
                }
            }
            wb.write(out);
        } finally {
            wb.close();
        }
    }


}
