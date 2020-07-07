package com.sugon.excel.entity;

import lombok.Data;

/**
 * @author litianfu
 * @version 1.0
 * @date 2020-07-07 17:02:59.151
 * @email 1035869369@qq.com
 * 要求该实体类是万能实体类
 */
@Data
public class ExcelEntity {

	private Long phone;
	private String sex;
	private String name;
	private Long age;
	private String hobby;

}
