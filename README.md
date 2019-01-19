# poi-plus

#### 介绍
excel工具

#### 使用说明
excel.json配置相关的读取参数，如果不填则使用系统默认参数

Excel 为主工具类

#####导出使用示例



1. 首先创建导出数据的model，在字段上添加@ExcelField注解，（如需要合并单元格则需要继承ExcelModel类）例如
		
		@ExcelField(columnName="ID",order = 1,isRowMerger = true)
    	private int id;

2. 注解参数主要包含

		/**
		 * 对应的列名（在excel中对应的列名）
		 * @return
		 */
		String[] columnName();
	
		/**
		 * 排序
		 * @return
		 */
		int order() default  0;
	
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

3. 如果有需要合并单元格，则需要将合并的字段和合并单元格的数量包装到MergeColumn对象里，然后放入MergeColumnList集合中。

4. 导入、导出时会根据字段上的注解读取属性的值，例如

		List<SysCommunityInfo> list = new ArrayList<>();
        SysCommunityInfo sysCommunityInfo = new SysCommunityInfo();
        sysCommunityInfo.setBuildingNo("1");
        sysCommunityInfo.setUnitNo("2");
        sysCommunityInfo.setDoorNo("3");
        sysCommunityInfo.setCommunityId(19);
        sysCommunityInfo.setId(1);
        sysCommunityInfo.setCreateTime(System.currentTimeMillis());
        SysCommunityInfo sysCommunityInfo1 = new SysCommunityInfo();
        sysCommunityInfo1.setBuildingNo("2");
        sysCommunityInfo1.setUnitNo("3");
        sysCommunityInfo1.setDoorNo("5");
        sysCommunityInfo1.setCommunityId(19);
        sysCommunityInfo1.setId(2);
        //合并单元格
        List<MergeColumn> mergeColumnList = new ArrayList<>();
		//id字段合并两个单元格
        MergeColumn mergeColumnId = new MergeColumn("id",2);
        MergeColumn mergeColumn = new MergeColumn("buildingNo",2);
        mergeColumnList.add(mergeColumn);
        mergeColumnList.add(mergeColumnId);
        sysCommunityInfo.setMergeColumnList(mergeColumnList);

        list.add(sysCommunityInfo);
        list.add(sysCommunityInfo1);
		//包装数据，Class为list中泛型的类型，list则为包装的数据
        Map<Class,List> map = new HashMap<>();
        map.put(SysCommunityInfo.class,list);
		//map的key为excel的sheet页名称，value为该sheet页中的数据
        Map<String,Map<Class,List>> mapData = new HashMap<>();

        mapData.put("sheet1",map);

		
        Excel.exportExcel(response,"hello.xls",mapData);



#####导人使用示例

1. 将注解上的columnName对应上excel中的title.@ExcelField(columnName="单元")

2. 第一个参数表示文件名，第二个是输入流，第三个则是导出的集合的类型
		
		FileInputStream is = new FileInputStream("F:\\test.xlsx");
		List<SysCommunityInfo> sysCommunityInfoList = Excel.readExcelToList("文件名",输入流,SysCommunityInfo.class);

。。。后续更新