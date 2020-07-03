package com.sugon.excel.controller;


import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;

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
     * 服务连通性测试
     * @return
     */
    @RequestMapping("/test")
    public String test(){
        System.out.println("test");
        return "test....";
    }

    /**
     * 对poi库进行测试
     * 读取excel文件
     * @return
     */
    @RequestMapping("/readExcelFile")
    public String readExcelFile() throws Exception{
        //读取excel文件
        File xlsxFile = new File("C:\\Users\\Blunt\\Desktop\\2020-7-3南宁学院案件明细.xlsx");
        //工作表
        Workbook sheets = WorkbookFactory.create(xlsxFile);
        //获取工作表个数
        int numberOfSheets = sheets.getNumberOfSheets();

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

            // 读数据。
            for (int row = 0; row < maxRowNum; row++) {
                Row r = sheet.getRow(row);
                for (int col = 0; col < cells; col++) {
                    System.out.print(r.getCell(col).toString()+" ");
                }

                // 换行。
                System.out.println();
            }


        }
        return null;
    }


}
