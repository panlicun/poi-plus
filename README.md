# poi-plus

#### 介绍
excel工具

----------
2.2版本更新内容

**1、去掉了excel.json配置文件的读取，改用通过ExcelConfig类进行相关的配置**
	
	导入功能
	readExcelToList() 去掉了excelFileName参数
	
	

**2、将原来的静态方法都修改成了对象的内部方法**

**3、ExcelStyle不再通过方法传参的方式传入，改为通过Excel属性的方式传入**

**4、新增了自定义单元格功能**

	使用示例
	
	//创建样式类，继承ExcelStyle，重写自己需要的样式
	ExcelStyle excelStyle = new MyStyle();
	//创建配置对象
    ExcelConfig excelConfig = new ExcelConfig();
	//创建自定义数据，构造方法，参数1为单元格要显示的文字，参数二为要合并的单元格
    ExcelCustomData excelCustomData = new ExcelCustomData("测试",new CellRangeAddress(0,0,0,4));
    ExcelCustomData excelCustomData2 = new ExcelCustomData("测试",new CellRangeAddress(1,1,0,4));
	//将自定义数据放入集合
    List<ExcelCustomData> excelCustomDataList = new ArrayList<>();
    excelCustomDataList.add(excelCustomData);
    excelCustomDataList.add(excelCustomData2);
	//将自定义数据放入配置文件
    excelConfig.setExcelCustomDatas(excelCustomDataList);
	//设置从哪行开始写入
    excelConfig.setWriteStartRow(2);
	//创建主类对象
    Excel excel = new Excel();
	//传入样式文件
    excel.setExcelStyle(excelStyle);
	//传入配置文件
    excel.setExcelConfig(excelConfig);
	//调用导出excel方法
    excel.exportExcel(response,"hello.xls",mapData);

**以上为新增功能**

----------
**以下使用说明为之前版本的使用说明，请按上面最新使用方式使用，如果不明，请参考之前的版本使用说明**


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
		List<SysCommunityInfo> sysCommunityInfoList = Excel.readExcelToList(输入流,SysCommunityInfo.class);



----------

###新增功能

可以自定义样式，使用如下

	继承ExcelStyle类，并且重写setTitleStyle或setDataStyle方法。然后将自定义样式的类的对象传到方法中

	ExcelStyle excelStyle = new MyStyle();
	Excel.exportExcel(response,"hello.xls",mapData,excelStyle);



列合并目前还没有完成，后续更新


。。。后续更新