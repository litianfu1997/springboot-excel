package com.sugon.excel.controller;


import com.sugon.excel.compiler.CompilerJob;
import com.sugon.excel.entity.EntityGenerator;
import com.sugon.excel.entity.*;
import com.sugon.excel.res.ResultEntity;
import com.sugon.excel.res.ResultEnum;
import com.sugon.excel.util.ChineseToSpell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.bind.annotation.*;

import java.io.File;
import java.io.FileInputStream;
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
     * 文件名
     */
    private String excelFile;


    /**
     * 按行存放的表格数据
     */
    private List<Map<String, Object>> resultData;


    /**
     * 服务连通性测试
     *
     * @return
     */
    @RequestMapping(value = "/test")
    public ResultEntity test() {
        System.out.println("test");
        System.out.println("test end123");


        return new ResultEntity(ResultEnum.SUCCESS, "entityMap");
    }


    /**
     * 通过行来读取excel表格数据
     *
     * @throws Exception
     */
    public List<Map<String, Object>> readExcelFileOnRow(String excelFile) {
        if (this.isExcelFile(excelFile).equals(false)) {
            logger.info("该文件不是excel文件");
            return null;
        }
        System.out.println(excelFile);
        if (excelFile == null) {
            logger.info("文件名为空！");
            return null;
        }
        this.excelFile = excelFile;
        Map<String, String> excelCellsType = this.getExcelCellsType();
        if (excelCellsType == null) {
            return null;
        }
        //动态创建ExcelEntity
        EntityGenerator.getInstance();
        //获取实体类生成器
        EntityGenerator entityGenerator = EntityGenerator.getEntityGenerator();
        entityGenerator.setMap(excelCellsType);
        //执行生成器
        entityGenerator.generator(ChineseToSpell.getFullSpell(excelFile));

        //对ExcelEntity进行编译
        compiler.compiler(ChineseToSpell.getFullSpell(excelFile));
        //读取excel文件
        File xlsxFile = new File(excelFilePath + excelFile);
        if (!xlsxFile.exists()) {
            logger.info("excel文件不存在！");
            return null;
        }
        String suffix = excelFile.substring(excelFile.lastIndexOf("."));
        //工作表
        Workbook sheets = null;
        try {
            if (".xls".equals(suffix) || ".csv".equals(suffix)) {
                sheets = new HSSFWorkbook(new FileInputStream(xlsxFile));
            } else {
                sheets = WorkbookFactory.create(xlsxFile);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        //存储封装好的excel所有数据
        resultData = new ArrayList<>();

        //只获取第一张表
        Sheet sheet = sheets.getSheetAt(0);
        // Excel第一行,就是表头
        Row temp = sheet.getRow(0);
        if (temp == null) {
            logger.info("该表为空");
            return null;
        }
        //获取该表行数
        int maxRowNum = sheet.getLastRowNum() + 1;
        //获取该表的列数
        int cells = temp.getPhysicalNumberOfCells();


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

        this.resultData = sheetList;
        return this.resultData;
    }


    /**
     * 获取对应excel文件所有行数据
     * @param excelFile excel文件名
     * @return
     */
    @GetMapping("/selectSheetAllRow")
    public ResultEntity selectSheetAllRow(String excelFile){
        List<Map<String, Object>> list = this.readExcelFileOnRow(excelFile);
        this.removeExcelEntity(excelFile);
        return new ResultEntity(ResultEnum.SUCCESS, list);
    }



    /**
     * 可以通过任何key进行组合查询
     * 对象的值全部传null，证明查询该表所有数据
     *
     * @param entityMapData 前端传递的查询字段
     * @return
     */
    @RequestMapping("/selectRowByAnyKeys")
    public ResultEntity selectRowByAnyKeys(@RequestBody Map<String, Object> entityMapData) {

        Map<String, String> excelFiledMap = new HashMap<>();

        Map<String, Object> entityMap = (Map<String, Object>) entityMapData.get("data");
        this.excelFile = (String) entityMapData.get("sheet");
        if (excelFile == null) {
            return new ResultEntity(0, "SYS_ERROR", "sheet为空");
        }

        //获取excel实体
        Object excelEntity = this.getExcelEntity();
        if (excelEntity == null) {
            return new ResultEntity(0, "SYS_ERROR", "该文件为空！");
        }
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
                excelFiledMap.put(entry.getKey(), key);
                //如果传入的key不存在
                if (!excelCellsType.containsKey(key)) {
                    return new ResultEntity(0, "SYS_ERROR", "key不存在！");
                }

                if (value == null) {
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
        List<Map<String, Object>> sheetData = resultData;

        int finalNotNullCount = notNullCount;
        //查询符合条件的数据
        Object finalExcelEntity = excelEntity;
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
                    if (e.get(key).equals(excelEntityMethods.get("get" + key.substring(0, 1).toUpperCase() + key.substring(1))
                            .invoke(finalExcelEntity))) {
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

        //删除实体类
        this.removeExcelEntity(excelFile);

        return new ResultEntity(ResultEnum.SUCCESS, result);
    }


    /**
     * 获取每个表头所对应列的名字与值类型
     *
     * @return
     */
    @GetMapping("/getExcelCellsType")
    public Map<String, String> getExcelCellsType() {
        //读取excel文件
        File xlsxFile = new File(excelFilePath + this.excelFile);
        String suffix = excelFile.substring(excelFile.lastIndexOf("."));
        //工作表
        Workbook sheets = null;
        try {
            if (".xls".equals(suffix) || ".csv".equals(suffix)) {
                sheets = new HSSFWorkbook(new FileInputStream(xlsxFile));
            } else {
                sheets = WorkbookFactory.create(xlsxFile);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        Map<String, String> map = null;
        //遍历表
        for (int i = 0; i < 1; i++) {
            //获取表
            Sheet sheet = sheets.getSheetAt(i);
            // Excel第一行,就是表头
            Row temp = sheet.getRow(0);
            if (temp == null) {
                logger.info("该表为空");
                return null;
            }
            //获取该表行数
            int maxRowNum = sheet.getLastRowNum() + 1;
            //获取该表的列数
            int cells = temp.getPhysicalNumberOfCells();

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

        List<Map<String, Object>> lists = this.readExcelFileOnRow(this.excelFile);
        if (lists == null) {
            return null;
        }
        Object excelEntity = null;

        try {
            //获取ExcelEntity
            Class<?> clazz = Class.forName("com.sugon.excel.entity.ExcelEntity"+
                    ChineseToSpell.getFullSpell(this.excelFile.split("\\.")[0]));

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

    /**
     * 判断该文件是否为excel/cvs文件
     *
     * @param file
     * @return
     */
    private Boolean isExcelFile(String file) {
        if (file.length() < 3 || file.indexOf(".") == -1) {
            return false;
        }
        String suffix = file.substring(file.lastIndexOf("."));
        if (".xlsx".equals(suffix) || ".xls".equals(suffix) || ".xlt".equals(suffix) || ".csv".equals(suffix)) {
            return true;
        }
        return false;


    }

    /**
     * 删除实体类
     * @param excelFile
     */
    private void removeExcelEntity(String excelFile){
        //删除实体类
        EntityGenerator.getInstance();
        EntityGenerator entityGenerator = EntityGenerator.getEntityGenerator();
        entityGenerator.removeEntity(ChineseToSpell.getFullSpell(excelFile));
    }


}

