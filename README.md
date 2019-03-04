# poi-plus

#### 介绍
excel工具

2.0版本

----------
2.0版本主要更新了导出功能，导入功能与原来一样使用

2.0新增加的注解属性 `isObject`

导出功能不用自己再组织数据格式，拿到要导出的集合以后，直接调用Excel.exportExcel()方法即可。直接将要导出的列上添加注解。如果有其他对象中的属性要导出，则在该对象上添加注解`@ExcelField(isObject=true)`，然后在该对象的属性上添加注解。如果合并单元格，则需要在合并单元格的字段上添加`isRowMerger`注解属性即可。不用再考虑哪一行合并，工具类会根据数据情况自动合并单元格。

#### 使用说明
excel.json配置相关的读取参数，如果不填则使用系统默认参数

Excel 为主工具类

#####导出使用示例



1. 在要导出的类的字段上添加@ExcelField注解

		用户类
	
		public class SysUser {

		    @ExcelField(columnName="ID",order = 1,isRowMerger = true)
		    private int id;
		    @ExcelField(columnName="姓名",order = 2,isRowMerger = true)
		    private String name;
			
			/**
			 * 如果该对象下有要导出的列，则需要添加如下注释
			 */
		    @ExcelField(isObject=true)
		    private List<SubjectType> subjectTypeList;
		
		
		    public int getId() {
		        return id;
		    }
		
		    public void setId(int id) {
		        this.id = id;
		    }
		
		    public String getName() {
		        return name;
		    }
		
		    public void setName(String name) {
		        this.name = name;
		    }
		
		    public List<SubjectType> getSubjectTypeList() {
		        return subjectTypeList;
		    }
		
		    public void setSubjectTypeList(List<SubjectType> subjectTypeList) {
		        this.subjectTypeList = subjectTypeList;
		    }
		}

		
		学科类型

		public class SubjectType {
		    private int id;
		
		    @ExcelField(columnName="学科类型",order = 3,isRowMerger = true)
		    private String subjectType;
		
		    @ExcelField(isObject=true)
		    private List<Subject> subjectList;
		
		    public int getId() {
		        return id;
		    }
		
		    public void setId(int id) {
		        this.id = id;
		    }
		
		    public String getSubjectType() {
		        return subjectType;
		    }
		
		    public void setSubjectType(String subjectType) {
		        this.subjectType = subjectType;
		    }
		
		    public List<Subject> getSubjectList() {
		        return subjectList;
		    }
		
		    public void setSubjectList(List<Subject> subjectList) {
		        this.subjectList = subjectList;
		    }
		}
		
		学科类

		public class Subject {
		    private int id;
		
		    @ExcelField(columnName="学科名称",order = 3)
		    private String subjectName;
		
		    @ExcelField(columnName="分数",order = 4)
		    private int score;
		
		    public Subject(int id, String subjectName, int score) {
		        this.id = id;
		        this.subjectName = subjectName;
		        this.score = score;
		    }
		
		    public int getId() {
		        return id;
		    }
		
		    public void setId(int id) {
		        this.id = id;
		    }
		
		    public String getSubjectName() {
		        return subjectName;
		    }
		
		    public void setSubjectName(String subjectName) {
		        this.subjectName = subjectName;
		    }
		
		    public int getScore() {
		        return score;
		    }
		
		    public void setScore(int score) {
		        this.score = score;
		    }
		}

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

		/**
		 * 是否是对象
		 */
		boolean isObject() default false;

3. 如果有需要合并单元格，需要在该列注解中添加属性 `isRowMerger = true` ，工具类会根绝数据格式自动合并单元格

4. 导入、导出时会根据字段上的注解读取属性的值，例如

		SysUser sysUser = new SysUser();
        sysUser.setId(1);
        sysUser.setName("张三");

        SubjectType subjectType = new SubjectType();
        subjectType.setId(1);
        subjectType.setSubjectType("理科");
        List<Subject> subjectList = new ArrayList<>();
        Subject subject1 = new Subject(1,"数学",96);
        Subject subject2 = new Subject(2,"物理",58);
        Subject subject3 = new Subject(3,"化学",90);
        subjectList.add(subject1);
        subjectList.add(subject2);
        subjectList.add(subject3);
        subjectType.setSubjectList(subjectList);

        SubjectType subjectType1 = new SubjectType();
        subjectType1.setId(2);
        subjectType1.setSubjectType("文科");
        List<Subject> subjectList1 = new ArrayList<>();
        Subject subject4 = new Subject(4,"政治",58);
        Subject subject5 = new Subject(5,"地理",47);
        subjectList1.add(subject4);
        subjectList1.add(subject5);
        subjectType1.setSubjectList(subjectList1);

        List<SubjectType> subjectTypeList = new ArrayList<>();
        subjectTypeList.add(subjectType);
        subjectTypeList.add(subjectType1);
        sysUser.setSubjectTypeList(subjectTypeList);




        SysUser sysUser1 = new SysUser();
        sysUser1.setId(2);
        sysUser1.setName("李四");

        subjectType = new SubjectType();
        subjectType.setId(1);
        subjectType.setSubjectType("理科");
        subjectList = new ArrayList<>();
        subject1 = new Subject(1,"数学",85);
        subject2 = new Subject(2,"物理",25);
        subject3 = new Subject(3,"化学",65);
        subjectList.add(subject1);
        subjectList.add(subject2);
        subjectList.add(subject3);
        subjectType.setSubjectList(subjectList);

        subjectType1 = new SubjectType();
        subjectType1.setId(2);
        subjectType1.setSubjectType("文科");
        subjectList1 = new ArrayList<>();
        subject4 = new Subject(4,"政治",76);
        subject5 = new Subject(5,"地理",68);
        subjectList1.add(subject4);
        subjectList1.add(subject5);
        subjectType1.setSubjectList(subjectList1);

        subjectTypeList = new ArrayList<>();
        subjectTypeList.add(subjectType);
        subjectTypeList.add(subjectType1);
        sysUser1.setSubjectTypeList(subjectTypeList);

        List<SysUser> sysUsers = new ArrayList<>();
        sysUsers.add(sysUser);
        sysUsers.add(sysUser1);

        Map<Class,List> map = new HashMap<>();
        map.put(SysUser.class,sysUsers);
        Map<String,Map<Class,List>> mapData = new HashMap<>();
        mapData.put("sheet1",map);
        ExportExcelUtils_new.exportExcel(response,"hello.xls",mapData);



#####导入使用示例

1. 将注解上的columnName对应上excel中的title.@ExcelField(columnName="单元")

2. 第一个参数表示文件名，第二个是输入流，第三个则是导出的集合的类型
		
		FileInputStream is = new FileInputStream("F:\\test.xlsx");
		List<SysCommunityInfo> sysCommunityInfoList = Excel.readExcelToList("文件名",输入流,SysCommunityInfo.class);



----------

###新增功能

可以自定义表头样式，使用如下

	继承ExcelStyle类，并且重写setTitleStyle方法。然后将自定义样式的类的对象传到方法中

	ExcelStyle excelStyle = new MyStyle();
	Excel.exportExcel(response,"hello.xls",mapData,excelStyle);



列合并目前还没有完成，后续更新


。。。后续更新