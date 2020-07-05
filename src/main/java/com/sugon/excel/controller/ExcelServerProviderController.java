package com.sugon.excel.controller;


import org.apache.poi.ss.usermodel.*;
import org.json.JSONObject;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * @author litianfu
 * @version 1.0
 * @date 2020/7/3 17:20
 * @email 1035869369@qq.com
 * 该类是对excel文件进行读取并提供对应api对excel文件进行查询
 */
@RestController
@RequestMapping("/excelServer")
public class ExcelServerProviderController {

    /**
     * excel文件路径
     */
    @Value("${excel.path}")
    private String excelFilePath;

    /**
     * 用于存放excel表格所有数据
     */
    private Map<String, Object> map;

    /**
     * 按行存放的表格数据
     */
    private List<List<Map<String,Object>>> resultData;

    /**
     * 若有多个工作表，使用list进行存储
     */
    private List<Map<String, Object>> mapList = new ArrayList<>();


    /**
     * 服务连通性测试
     *
     * @return
     */
    @RequestMapping("/test")
    public String test() {
        System.out.println("test");
        return "test....";
    }

    /**
     * 对poi库进行测试
     * 读取excel文件,按列存储
     *
     * @return
     */
    @RequestMapping("/readExcelFileOnColumn")
    public void readExcelFileOnColumn() throws Exception {


        //读取excel文件
        File xlsxFile = new File(excelFilePath);
        //工作表
        Workbook sheets = WorkbookFactory.create(xlsxFile);
        //获取工作表个数
        int numberOfSheets = sheets.getNumberOfSheets();


        //存放第几张工作表的表头信息
        List<List<String>> workbooks = new ArrayList<>();
        //遍历表
        for (int i = 0; i < numberOfSheets; i++) {
            map = new HashMap<>();
            //获取表
            Sheet sheet = sheets.getSheetAt(i);

            //获取该表行数
            int maxRowNum = sheet.getLastRowNum() + 1;
            // Excel第一行,就是表头
            Row temp = sheet.getRow(0);
            if (temp == null) {
                continue;
            }
            //获取不为空的的列个数
            int cells = temp.getPhysicalNumberOfCells();


            //存放所有表头信息
            ArrayList<String> keys = new ArrayList<>();
            //读取数据
            for (int col = 0; col < cells; col++) {
                List<String> list = new ArrayList<>();
                for (int row = 1; row < maxRowNum; row++) {
                    //获取对应列的所有数据
                    Cell cell = sheet.getRow(row).getCell(col);
                    //将数据存入list
                    list.add(cell.toString());
                }
                //存放表头信息
                keys.add(sheet.getRow(0).getCell(col).toString());

                //表头作为map的键，值为list
                map.put(sheet.getRow(0).getCell(col).toString(), list);
            }
            workbooks.add(keys);
            //存放所有表头信息
            map.put("key", workbooks);

            //将不同工作表的数据存储到list中
            mapList.add(map);
            //清除map缓存
//            System.out.println(map);
        }

    }

    /**
     * 通过行来存储excel表格数据
     *
     * @throws Exception
     */
    @RequestMapping("/readExcelFileOnRow")
    public List<List<Map<String,Object>>> readExcelFileOnRow() throws Exception {
        //读取excel文件
        File xlsxFile = new File(excelFilePath);
        //工作表
        Workbook sheets = WorkbookFactory.create(xlsxFile);
        //获取工作表个数
        int numberOfSheets = sheets.getNumberOfSheets();

        resultData = new ArrayList<>();


        //遍历表
        for (int i = 0; i < numberOfSheets; i++) {
            //获取表
            Sheet sheet = sheets.getSheetAt(i);
            // Excel第一行,就是表头
            Row temp = sheet.getRow(0);
            //获取该表行数
            int maxRowNum = sheet.getLastRowNum() + 1;
            //获取该表的列数
            int cells = temp.getPhysicalNumberOfCells();

            if (temp == null) {
                continue;
            }
            List<Map<String,Object>> sheetList = new ArrayList<>();



            //遍历整个表格
            for (int row = 1; row < maxRowNum; row++) {
                //获取行
                Row sheetRow = sheet.getRow(row);
                Map<String, Object> map= new HashMap<>();
                for (int col = 0; col < cells; col++) {
                    //将每一行数据存储成一个json对象
                    map.put(sheet.getRow(0).getCell(col).toString()
                            , sheetRow.getCell(col).toString());
                }
                sheetList.add(map);


            }

            resultData.add(sheetList);
        }
        System.out.println(resultData);
        return resultData;
    }

    /**
     * 通过关键信息查询表格数据
     * （行存储方式）
     *
     * @param key
     * @param val
     */
    @RequestMapping("selectRowByKeyOnRow")
    public List<Map<String,Object>> selectRowByKeyOnRow(@RequestParam("key") String key,
                                    @RequestParam("val") String val,
                                    @RequestParam("workbook")Integer workbook) throws Exception {
        this.readExcelFileOnRow();

        //获取哪一张工作表的数据
        List<Map<String, Object>> sheetList = (List<Map<String, Object>>) resultData.get(workbook);

        List<Map<String, Object>> result = sheetList.stream()
                .filter(e -> e.get(key).equals(val))
                .collect(Collectors.toList());
        System.out.println(result);
        return result;

    }


    /**
     * 通过某个表头的数据查询数据
     * （列存储方式）
     *
     * @param workbook 第几张工作表（从0开始）
     * @param key      某个表头的值
     * @param val      表头对应的列中的某个值
     * @return 返回满足查找条件的行数据
     */
    @RequestMapping("/selectRowByKeyOnColumn")
    public List<Map<String, Object>> selectRowByKeyOnColumn(@RequestParam("key") String key,
                                                            @RequestParam("val") String val,
                                                            @RequestParam("workbook") Integer workbook) throws Exception {
        this.readExcelFileOnColumn();
        //获第几张工作表的整张表信息
        Map<String, Object> map = mapList.get(workbook);
        //通过key获取哪一列的数据
        List<String> colList = (List<String>) map.get(key);
        //用于存放某列对应值的下标
        List<Integer> rowIndex = new ArrayList<>();
        //查询符合条件的行
        if (colList == null) {
            return null;
        }
        for (int i = 0; i < colList.size(); i++) {
            if (colList.get(i).equals(val)) {
                rowIndex.add(i);
            }
        }

        List<List<String>> keyList = (List<List<String>>) map.get("key");
        System.out.println("******************************");
        //获取到表头，也就是map的key
        List<String> list = keyList.get(workbook);
        //存放查询结果集
        List<Map<String, Object>> resultList = new ArrayList<>();
        //存放一行数据
        Map<String, Object> rowMap = new HashMap<>();
        int index = 0;
        while (index < rowIndex.size()) {
            //从第一列开始查询对应数据
            for (int i = 0; i < list.size(); i++) {
                //获取对应表头下的所有行数据
                List subColList = (List) map.get(list.get(i));
                for (int j = 0; j < subColList.size(); j++) {
                    if (j == rowIndex.get(index)) {
                        rowMap.put(list.get(i), subColList.get(j));
                    }
                }

            }
            resultList.add(rowMap);
            index++;
        }
        System.out.println(resultList);
        return resultList;
    }


}
