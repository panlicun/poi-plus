package com.plc.annotation;

import java.lang.annotation.*;

@Target({ElementType.METHOD,ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelTitle {

	/**
	 * 对应的列名（在excel中对应的列名）
	 * @return
	 */
	String[] value();
	

}
