package com.sugon.excel.controller;


import com.sugon.excel.compiler.CompilerJob;
import com.sugon.excel.entity.EntityGenerator;
import com.sugon.excel.res.ResultEntity;
import com.sugon.excel.res.ResultEnum;
import com.sugon.excel.util.ChineseToSpell;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.bind.annotation.*;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.time.LocalDate;
import java.util.*;
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

    private final static Logger logger = LoggerFactory.getLogger(ExcelServerProviderController.class);


    /**
     * excel文件路径
     */
    @Value("${excel.path}")
    private String excelFilePath;

    /**
     * 编译器
     */
    @Autowired
    private CompilerJob compiler;


    /**
     * 用于存放excel表格所有数据
     */
    private Map<String, Object> map;

    /**
     * 按行存放的表格数据
     */
    private List<List<Map<String, Object>>> resultData;

    /**
     * 若有多个工作表，使用list进行存储
     */
    private List<Map<String, Object>> mapList = new ArrayList<>();


    /**
     * 服务连通性测试
     *
     * @return
     */
    @RequestMapping(value = "/test")
    public ResultEntity test(@RequestParam Map<String, Object> entityMap) {
        System.out.println("test");
//        Object excelEntity = this.getExcelEntity();
//        Map<String, String> excelCellsType = this.getExcelCellsType();
//        //获取实体类的所有getter和setter方法
//        Map<String, Method> excelEntityMethods = this.getExcelEntityMethods(excelEntity);
//        try {
//
//            //遍历map
//            Iterator<Map.Entry<String, Object>> entries = entityMap.entrySet().iterator();
//            while (entries.hasNext()) {
//                Map.Entry<String, Object> entry = entries.next();
//                String key = entry.getKey();
//                Object value = entry.getValue();
//                switch (excelCellsType.get(key)) {
//                    case "String":
//                        excelEntityMethods.get("set" + key.substring(0, 1).toUpperCase() + key.substring(1)).invoke(excelEntity
//                                , value.toString());
//                        break;
//                    case "Long":
//                        excelEntityMethods.get("set" + key.substring(0, 1).toUpperCase() + key.substring(1)).invoke(excelEntity
//                                , Long.parseLong(value.toString()));
//                        break;
//                    case "Boolean":
//                        excelEntityMethods.get("set" + key.substring(0, 1).toUpperCase() + key.substring(1)).invoke(excelEntity
//                                , Boolean.parseBoolean(value.toString()));
//                        break;
//                    case "Date":
//                        excelEntityMethods.get("set" + key.substring(0, 1).toUpperCase() + key.substring(1)).invoke(excelEntity
//                                , LocalDate.parse(value.toString()));
//                        break;
//                    default:
//                        excelEntityMethods.get("set" + key.substring(0, 1).toUpperCase() + key.substring(1)).invoke(excelEntity
//                                , value.toString());
//                        break;
//                }
//
//
//            }
//        } catch (IllegalAccessException e) {
//            e.printStackTrace();
//        } catch (InvocationTargetException e) {
//            e.printStackTrace();
//        }
//        System.out.println(excelEntity);
        return new ResultEntity(ResultEnum.SUCCESS,entityMap);
    }

    /**
     * 对poi库进行测试
     * 读取excel文件,按列存储
     *
     * @return
     */
    @GetMapping("/readExcelFileOnColumn")
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
    @GetMapping("/readExcelFileOnRow")
    public List<List<Map<String, Object>>> readExcelFileOnRow() {

        //动态创建ExcelEntity
        EntityGenerator.getInstance(this.getExcelCellsType());
        //获取实体类生成器
        EntityGenerator entityGenerator = EntityGenerator.getEntityGenerator();
        //执行生成器
        entityGenerator.generator();

        //对ExcelEntity进行编译
        compiler.compiler();
        //读取excel文件
        File xlsxFile = new File(excelFilePath);
        if (!xlsxFile.exists()){
            logger.info("excel文件不存在！");
            return null;
        }

        //工作表
        Workbook sheets = null;
        try {
            sheets = WorkbookFactory.create(xlsxFile);
        } catch (IOException e) {
            e.printStackTrace();
        }
        //获取工作表个数
        int numberOfSheets = sheets.getNumberOfSheets();

        //存储封装好的excel所有数据
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
            //存储不同工作表的数据
            List<Map<String, Object>> sheetList = new ArrayList<>();


            //遍历整个表格
            for (int row = 1; row < maxRowNum; row++) {
                //获取行
                Row sheetRow = sheet.getRow(row);
                //将每一行数据封装到map中
                Map<String, Object> map = new HashMap<>();
                //将每一行数据存储成一个json对象
                for (int col = 0; col < cells; col++) {
                    //如果遇到空的单元格，进行跳过处理
                    if (sheetRow.getCell(col) == null) {
                        map.put(ChineseToSpell.getPingYin(sheet.getRow(0).getCell(col).toString()), "null");
                        continue;
                    }
                    switch (sheetRow.getCell(col).getCellType()) {
                        case NUMERIC:
//                            Integer.valueOf((int)sheetRow.getCell(col).getNumericCellValue())
                            double doubleVal = sheetRow.getCell(col).getNumericCellValue();
                            long longVal = Math.round(doubleVal);
                            if (Double.parseDouble(longVal + ".0") == doubleVal) {
                                map.put(ChineseToSpell.getPingYin(sheet.getRow(0).getCell(col).toString()), longVal);
                            } else {
                                map.put(ChineseToSpell.getPingYin(sheet.getRow(0).getCell(col).toString()), doubleVal);
                            }

                            break;
                        case STRING:
                            map.put(ChineseToSpell.getPingYin(sheet.getRow(0).getCell(col).toString())
                                    , sheetRow.getCell(col).toString());
                            break;
                        case _NONE:
                            map.put(ChineseToSpell.getPingYin(sheet.getRow(0).getCell(col).toString()), "");
                            break;
                        case BLANK:
                            map.put(ChineseToSpell.getPingYin(sheet.getRow(0).getCell(col).toString()), "");
                            break;
                        case BOOLEAN:
                            map.put(ChineseToSpell.getPingYin(sheet.getRow(0).getCell(col).toString())
                                    , Boolean.valueOf(sheetRow.getCell(col).getBooleanCellValue()));
                            break;
                        case ERROR:
                            map.put(ChineseToSpell.getPingYin(sheet.getRow(0).getCell(col).toString()), "error");
                            break;
                        default:
                            map.put(ChineseToSpell.getPingYin(sheet.getRow(0).getCell(col).toString()), sheetRow.getCell(col).toString());
                    }


                }
                sheetList.add(map);


            }

            resultData.add(sheetList);
        }
        System.out.println(resultData);
        return resultData;
    }

    /**
     * 通过关键信息精准查询表格数据
     * （行存储方式）
     *
     * @param key       某个表头的值
     * @param val       表头对应的列中的某个值
     * @param workbook  第几张工作表
     * @param sort      排序方式（asc 升序、 desc 降序）
     * @param sortedKey 要排序的字段
     */
    @GetMapping("selectRowByKeyOnRow")
    public List<Map<String, Object>> selectRowByKeyOnRow(@RequestParam("key") String key,
                                                         @RequestParam("val") String val,
                                                         @RequestParam("workbook") Integer workbook,
                                                         @RequestParam(value = "sortedKey", required = false) String sortedKey,
                                                         @RequestParam(value = "sort", required = false) String sort) {
        //读取excel文件
        this.readExcelFileOnRow();

        //获取哪一张工作表的数据
        List<Map<String, Object>> sheetList = (List<Map<String, Object>>) resultData.get(workbook);


        List<Map<String, Object>> result = sheetList.stream()
                .filter(e -> e.get(key).equals(val))
                //自定义排序
                .sorted((o1, o2) -> {
                    //如果没有设置排序选项，默认为升序排序
                    if (!(sort == null || "".equals(sort) || sortedKey == null || "".equals(sortedKey))) {
                        //升序
                        if ("asc".equals(sort)) {
                            return o1.get(sortedKey).toString().compareTo(o2.get(sortedKey).toString());
                        }
                        //降序
                        if ("desc".equals(sort)) {
                            return o2.get(sortedKey).toString().compareTo(o1.get(sortedKey).toString());
                        }
                    }
                    return 1;
                })
                .collect(Collectors.toList());
        return result;

    }

    /**
     * 可以通过任何key进行组合查询
     * 不传任何参数，证明查询该表所有数据
     *
     * @param entityMapData 前端传递的查询字段
     * @return
     */
    @RequestMapping("/selectRowByAnyKeys")
    public ResultEntity selectRowByAnyKeys(@RequestBody Map<String, Object> entityMapData) {

        Map<String,String> excelFiledMap = new HashMap<>();

        Map<String, Object> entityMap = (Map<String, Object>) entityMapData.get("data");

        //获取excel实体
        Object excelEntity = this.getExcelEntity();
        //获取excel表格数据类型
        Map<String, String> excelCellsType = this.getExcelCellsType();
        //获取实体类的所有getter和setter方法
        Map<String, Method> excelEntityMethods = this.getExcelEntityMethods(excelEntity);

        //存放entityMap不为空的元素
        int notNullCount = 0;
        try {

            //遍历前端传递过来的entityMap
            Iterator<Map.Entry<String, Object>> entries = entityMap.entrySet().iterator();
            while (entries.hasNext()) {

                Map.Entry<String, Object> entry = entries.next();

                //将中文字段转换为拼音
                String key = ChineseToSpell.getPingYin(entry.getKey());
                Object value = entry.getValue();
                //excel表格表头字段
                excelFiledMap.put(entry.getKey(),key);
                //如果传入的key不存在
                if (!excelCellsType.containsKey(key)) {
                    return new ResultEntity(0,"SYS_ERROR","key不存在！");
                }

                if (value == null){
                    continue;
                }
                //计算entityMap不为空的元素
                notNullCount++;


                //封装实体类
                switch (excelCellsType.get(key)) {
                    case "String":
                        excelEntityMethods.get("set" + key.substring(0, 1).toUpperCase() + key.substring(1)).invoke(excelEntity
                                , value.toString());
                        break;
                    case "Long":
                        excelEntityMethods.get("set" + key.substring(0, 1).toUpperCase() + key.substring(1)).invoke(excelEntity
                                , Long.parseLong(value.toString()));
                        break;
                    case "Boolean":
                        excelEntityMethods.get("set" + key.substring(0, 1).toUpperCase() + key.substring(1)).invoke(excelEntity
                                , Boolean.parseBoolean(value.toString()));
                        break;
                    case "Date":
                        excelEntityMethods.get("set" + key.substring(0, 1).toUpperCase() + key.substring(1)).invoke(excelEntity
                                , LocalDate.parse(value.toString()));
                        break;
                    default:
                        excelEntityMethods.get("set" + key.substring(0, 1).toUpperCase() + key.substring(1)).invoke(excelEntity
                                , value.toString());
                        break;
                }


            }
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (InvocationTargetException e) {
            e.printStackTrace();
        }

        //获取第一张工作表
        List<Map<String, Object>> sheetData = resultData.get(0);

        int finalNotNullCount = notNullCount;
        //查询符合条件的数据
        List<Map<String, Object>> result = sheetData.stream().filter(e -> {
            //遍历map
            Iterator<Map.Entry<String, String>> keyEntries = excelCellsType.entrySet().iterator();
            List<Boolean> booleanList = new ArrayList<>();
            while (keyEntries.hasNext()) {
                Map.Entry<String, String> entry = keyEntries.next();
                String key = entry.getKey();
                Object value = entry.getValue();
                try {
                    //将满足条件的数据添加标志
                    if (e.get(key).equals(excelEntityMethods.get("get" + key.substring(0, 1).toUpperCase() + key.substring(1)).invoke(excelEntity))) {
                        booleanList.add(true);
                    }
                } catch (IllegalAccessException ex) {
                    ex.printStackTrace();
                } catch (InvocationTargetException ex) {
                    ex.printStackTrace();
                }
            }
            //判断标志数量是否等于传参个数
            return booleanList.size() == finalNotNullCount ? true : false;
        }).collect(Collectors.toList());

        System.out.println("result = " + result);

        return new ResultEntity(ResultEnum.SUCCESS,result);
    }

    /**
     * 通过关键信息模糊查询表格数据
     * （行存储方式）
     *
     * @param key       某个表头的值
     * @param likeVal   表头对应的列中的某个模糊查询的值
     * @param workbook  第几张工作表
     * @param sort      排序方式（asc 升序、 desc 降序）
     * @param sortedKey 要排序的字段
     * @return
     */
    @GetMapping("/selectRowByLikeKeyOnRow")
    public List<Map<String, Object>> selectRowByLikeKeyOnRow(@RequestParam(value = "key") String key,
                                                             @RequestParam("likeVal") String likeVal,
                                                             @RequestParam("workbook") Integer workbook,
                                                             @RequestParam(value = "sortedKey", required = false) String sortedKey,
                                                             @RequestParam(value = "sort", required = false) String sort) {
        //读取excel文件
        this.readExcelFileOnRow();

        //获取哪一张工作表的数据
        List<Map<String, Object>> sheetList = (List<Map<String, Object>>) resultData.get(workbook);

        List<Map<String, Object>> result = sheetList.stream()
                .filter(e -> e.get(key).toString().indexOf(likeVal) != -1)
                //自定义排序
                .sorted((o1, o2) -> {
                    //如果没有设置排序选项，默认为升序排序
                    if (!(sort == null || "".equals(sort) || sortedKey == null || "".equals(sortedKey))) {
                        //升序
                        if ("asc".equals(sort)) {
                            return o1.get(sortedKey).toString().compareTo(o2.get(sortedKey).toString());
                        }
                        //降序
                        if ("desc".equals(sort)) {
                            return o2.get(sortedKey).toString().compareTo(o1.get(sortedKey).toString());
                        }
                    }
                    return 1;
                })
                .collect(Collectors.toList());


        return result;
    }


    /**
     * 按某个表头进行分组
     *
     * @param key       某个表头的值
     * @param workbook  第几张工作表
     * @param sortedKey 排序方式（asc 升序、 desc 降序）
     * @param sort      要排序的字段
     * @return
     */
    public List<Map<String, Object>> groupByKey(@RequestParam("key") String key,
                                                @RequestParam("workbook") Integer workbook,
                                                @RequestParam(value = "sortedKey", required = false) String sortedKey,
                                                @RequestParam(value = "sort", required = false) String sort) {

        return null;
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
    @GetMapping("/selectRowByKeyOnColumn")
    public List<Map<String, Object>> selectRowByKeyOnColumn(@RequestParam("key") String key,
                                                            @RequestParam("val") String val,
                                                            @RequestParam("workbook") Integer workbook) throws Exception {
        //读取excel文件
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


    /**
     * 获取每个表头所对应列的名字与值类型
     *
     * @return
     */
    @GetMapping("/getExcelCellsType")
    public Map<String, String> getExcelCellsType() {
        //读取excel文件
        File xlsxFile = new File(excelFilePath);
        //工作表
        Workbook sheets = null;
        try {
            sheets = WorkbookFactory.create(xlsxFile);
        } catch (IOException e) {
            e.printStackTrace();
        }
        //获取工作表个数
        int numberOfSheets = sheets.getNumberOfSheets();
        Map<String, String> map = null;
        //遍历表
        for (int i = 0; i < 1; i++) {
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

            //遍历整个表格
            for (int row = 1; row < 2; row++) {
                //获取行
                Row sheetRow = sheet.getRow(row);
                //将每一行数据封装到map中
                map = new HashMap<>();
                for (int col = 0; col < cells; col++) {
                    switch (sheetRow.getCell(col).getCellType().toString()) {
                        case "STRING":
                            map.put(ChineseToSpell.getPingYin(sheet.getRow(0).getCell(col).toString()), "String");
                            break;
                        case "NUMERIC":
                            map.put(ChineseToSpell.getPingYin(sheet.getRow(0).getCell(col).toString()), "Long");
                            break;
                        case "BOOL":
                            map.put(ChineseToSpell.getPingYin(sheet.getRow(0).getCell(col).toString()), "Boolean");
                            break;
                        case "DATE":
                            map.put(ChineseToSpell.getPingYin(sheet.getRow(0).getCell(col).toString()), "Date");
                            break;
                        default:
                            map.put(ChineseToSpell.getPingYin(sheet.getRow(0).getCell(col).toString()), "String");
                            break;
                    }
                    System.out.print(sheet.getRow(0).getCell(col).toString()
                            + "=" + sheetRow.getCell(col).getCellType().toString() + " ");
                }


            }

        }
        return map;


    }


    /**
     * 反射获取ExcelEntity实例
     *
     * @return
     */
    public Object getExcelEntity() {

        this.readExcelFileOnRow();
        Object excelEntity = null;

        try {
            //获取ExcelEntity
            Class<?> clazz = Class.forName("com.sugon.excel.entity.ExcelEntity");
            excelEntity = clazz.newInstance();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return excelEntity;
    }

    /**
     * 获取ExcelEntity的所有getter和setter方法
     * 调用方式 map.get(方法名)
     *
     * @return
     */
    public Map<String, Method> getExcelEntityMethods(Object excelEntity) {

        Map<String, Method> methodsMap = null;
        try {

            //存放字段的map
            Map<String, String> fieldMap = this.getExcelCellsType();

            //存放方法的map
            methodsMap = new HashMap<>();
            Method[] declaredMethods = excelEntity.getClass().getDeclaredMethods();
            for (Method declaredMethod : declaredMethods) {
                System.out.println("declaredMethod.getName() = " + declaredMethod.getName());
            }
            //对fieldMap进行遍历
            Iterator<Map.Entry<String, String>> entries = fieldMap.entrySet().iterator();
            while (entries.hasNext()) {
                Map.Entry<String, String> entry = entries.next();
                String key = entry.getKey();
                String value = entry.getValue();
                //setter方法
                Method setMethod = null;
                //getter方法
                Method getMethod = null;
                switch (value) {
                    case "String":
                        setMethod = excelEntity.getClass().getDeclaredMethod("set" + key.substring(0, 1).toUpperCase() + key.substring(1), String.class);
                        break;
                    case "Long":
                        setMethod = excelEntity.getClass().getDeclaredMethod("set" + key.substring(0, 1).toUpperCase() + key.substring(1), Long.class);
                        break;
                    case "Date":
                        setMethod = excelEntity.getClass().getDeclaredMethod("set" + key.substring(0, 1).toUpperCase() + key.substring(1), Date.class);
                        break;
                    case "Boolean":
                        setMethod = excelEntity.getClass().getDeclaredMethod("set" + key.substring(0, 1).toUpperCase() + key.substring(1), Boolean.class);
                        break;
                    default:
                        setMethod = excelEntity.getClass().getDeclaredMethod("set" + key.substring(0, 1).toUpperCase() + key.substring(1), String.class);
                        break;

                }
                //将字符串进行组合，如getAge
                getMethod = excelEntity.getClass().getDeclaredMethod("get" + key.substring(0, 1).toUpperCase() + key.substring(1), null);
                methodsMap.put("get" + key.substring(0, 1).toUpperCase() + key.substring(1), getMethod);
                methodsMap.put("set" + key.substring(0, 1).toUpperCase() + key.substring(1), setMethod);

            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return methodsMap;
    }


}
