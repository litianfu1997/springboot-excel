package com.sugon.excel.controller;


import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

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
    private Map<String, List<String>> map;

    /**
     * 若有多个工作表，使用list进行存储
     */
    private List<Map<String, List<String>>> mapList;


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
     * 读取excel文件
     *
     * @return
     */
    @RequestMapping("/readExcelFile")
    public String readExcelFile() throws Exception {


        //读取excel文件
        File xlsxFile = new File(excelFilePath);
        //工作表
        Workbook sheets = WorkbookFactory.create(xlsxFile);
        //获取工作表个数
        int numberOfSheets = sheets.getNumberOfSheets();

         map = new HashMap<>();

        //遍历表
        for (int i = 0; i < numberOfSheets; i++) {
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

            //读取数据
            for (int col = 0; col < cells; col++) {
                List<String> list = new ArrayList<>();
                for (int row = 1; row < maxRowNum; row++) {
                    //获取对应列的所有数据
                    Cell cell = sheet.getRow(row).getCell(col);
                    //将数据存入list
                    list.add(cell.toString());
                }
                //表头作为map的键，值为list
                map.put(sheet.getRow(0).getCell(col).toString(), list);
            }
            System.out.println(map);
            //将不同工作表的数据存储到list中
            mapList = new ArrayList<>();
            mapList.add(map);
            //清除map缓存
            map.clear();
        }

        return null;
    }


}
