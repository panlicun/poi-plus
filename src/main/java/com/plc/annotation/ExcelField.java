package com.plc.annotation;

import java.lang.annotation.*;

@Target({ElementType.METHOD,ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelField {

	/**
	 * 对应的列名（在excel中对应的列名）
	 * @return
	 */
	String[] columnName() default {};

	/**
	 * 排序
	 * @return
	 */
	int order() default 100;

	/**
	 * 是否是日期
	 *
	 * @return
	 */
	boolean isDate() default false;

	/**
	 * 格式
	 *
	 * @return
	 */
	String fomat() default "yyyy/MM/dd";

	/**
	 * 是否是数字
	 *
	 * @return
	 */
	boolean isNum() default false;//data是否为数值型

	/**
	 * 是否是整型
	 *
	 * @return
	 */
	boolean isInteger() default false;//data是否为整型


	/**
	 * 是否列合并
	 *
	 * @return
	 */
	boolean isColMerger() default false;//是否横向合并单元格

	/**
	 * 是否行合并
	 *
	 * @return
	 */
	boolean isRowMerger() default false;//是否纵向合并单元格

	/**
	 * 列宽
	 */
	int columnWidth() default 0 * 256;

	boolean isObject() default false;


}
