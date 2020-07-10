package com.sugon.excel.entity;

import org.apache.xmlbeans.impl.xb.xsdschema.Public;

import java.io.*;
import java.time.LocalDate;
import java.time.LocalTime;
import java.util.Iterator;
import java.util.Map;

/**
 * @author litianfu
 * @version 1.0
 * @date 2020/7/6 10:55
 * @email 1035869369@qq.com
 * 尝试
 */
public class EntityGenerator {
    private static volatile EntityGenerator entityGenerator;

    private static Map<String, String> map;

    private EntityGenerator() {

    }

    /**
     * 实例化生成器
     */
    public static void getInstance() {
        if (entityGenerator == null) {
            synchronized (EntityGenerator.class) {
                if (entityGenerator == null) {
                    entityGenerator = new EntityGenerator();
                }
            }
        }
    }

    public  void setMap(Map<String, String> map) {
        EntityGenerator.map = map;
    }

    public static EntityGenerator getEntityGenerator() {
        return entityGenerator;
    }

    /**
     * 生成excel对应的实体类
     */
    public void generator(String fileName) {

        String name = fileName.split("\\.")[0];


        //编写实体类文件
        try {
            File file = new File(System.getProperty("user.dir") + "\\src\\main\\java\\com\\sugon\\excel\\entity" + "\\ExcelEntity"+name+".java");
            if (!file.exists()) {
                file.createNewFile();
            }
            OutputStream outputStream = new FileOutputStream(System.getProperty("user.dir") + "\\src\\main\\java\\com\\sugon\\excel\\entity" + "\\ExcelEntity"+name+".java");
            outputStream.write((
                    "package com.sugon.excel.entity;\n" +
                            "\n" +
                            "import lombok.Data;\n" +
                            "\n" +
                            "/**\n" +
                            " * @author litianfu\n" +
                            " * @version 1.0\n" +
                            " * @date " + LocalDate.now() + " " + LocalTime.now() + "\n" +
                            " * @email 1035869369@qq.com\n" +
                            " * 要求该实体类是万能实体类\n" +
                            " */\n" +
                            "@Data\n" +
                            "public class ExcelEntity"+name+" {\n" +
                            "\n"
            ).getBytes());
            //遍历map
            Iterator<Map.Entry<String, String>> entries = map.entrySet().iterator();
            while (entries.hasNext()) {
                Map.Entry<String, String> entry = entries.next();
                String key = entry.getKey();
                String value = entry.getValue();
                outputStream.write(("\tprivate " + value + " " + key + ";\n").getBytes());
            }

            outputStream.write("\n}\n".getBytes());
            outputStream.flush();
            outputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    public void removeEntity(String fileName){
        String name = fileName.split("\\.")[0];
        File file = new File(System.getProperty("user.dir") + "\\src\\main\\java\\com\\sugon\\excel\\entity" + "\\ExcelEntity"+name+".java");
        File classFile = new File(System.getProperty("user.dir") + "\\target\\classes\\com\\sugon\\excel\\entity\\ExcelEntity"+name+".class");
        // 路径为文件且不为空则进行删除
        if (file.isFile() && file.exists()) {
            boolean delete = file.delete();

        }
        if (classFile.isFile()&& classFile.exists()){
            boolean delete = classFile.delete();
        }
    }

}
