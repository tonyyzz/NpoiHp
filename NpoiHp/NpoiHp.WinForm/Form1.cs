using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NpoiHp.WinForm
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}



		//文档：http://blog.csdn.net/pan_junbiao/article/details/39717443

		private void btnGenerate_Click(object sender, EventArgs e)
		{

			HSSFWorkbook hssfworkbook = new HSSFWorkbook();



			//1
			#region 初始化
			////创建xls
			//HSSFWorkbook hssfworkbook = new HSSFWorkbook();
			////创建工作表
			//var sheet = hssfworkbook.CreateSheet("newsheet");

			////生成文件
			//WriteXlsToFile(hssfworkbook, @"test.xls");
			#endregion

			//2
			#region 创建DocumentSummaryInformation和SummaryInformation
			//DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
			//dsi.Company = "NPOI Team";
			//SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
			//si.Subject = "NPOI SDK Example";
			//HSSFWorkbook hssfworkbook = new HSSFWorkbook()
			//{
			//	DocumentSummaryInformation = dsi,
			//	SummaryInformation = si
			//}; 
			#endregion

			//3
			//创建工作表
			//var sheet = hssfworkbook.CreateSheet("sheet1");
			////创建行
			//var row = sheet.CreateRow(0);
			////创建单元格
			//var cell = row.CreateCell(0);
			//赋值
			//cell.SetCellValue(1);

			//var cell1 = sheet.GetRow(0).GetCell(0);
			//var v = cell.NumericCellValue;

			//var cIndex = cell.ColumnIndex;
			//var rIndex = cell.RowIndex;

			#region 创建批注
			/*
			    参数 说明
				dx1 第1个单元格中x轴的偏移量
				dy1 第1个单元格中y轴的偏移量
				dx2 第2个单元格中x轴的偏移量
				dy2 第2个单元格中y轴的偏移量
				col1 第1个单元格的列号
				row1 第1个单元格的行号
				col2 第2个单元格的列号
				row2 第2个单元格的行号
			 */
			//var patr = sheet.CreateDrawingPatriarch() as HSSFPatriarch;
			////创建批注
			//var comment1 = patr?.CreateComment(new HSSFClientAnchor(0, 0, 0, 0, cIndex + 1, rIndex + 1, cIndex + 5, rIndex + 5));
			//comment1.String = new HSSFRichTextString("Hello World");
			//comment1.Author = "NPOI Team";
			//comment1.Visible = true;
			//cell.CellComment = comment1; 
			#endregion


			#region 创建页眉页脚
			//set headertext
			//sheet.Header.Center = "This is a test sheet";
			////set footertext
			//sheet.Footer.Left = "Copyright NPOI Team";
			//sheet.Footer.Right = "created by Tony Qu（瞿杰）&D"; 
			#endregion

			#region 设置日期格式
			//cell.SetCellValue(DateTime.Now);

			//HSSFCellStyle cellStyle = hssfworkbook.CreateCellStyle() as HSSFCellStyle;
			//HSSFDataFormat format = hssfworkbook.CreateDataFormat() as HSSFDataFormat;
			//cellStyle.DataFormat = format.GetFormat("yyyy年MM月dd日 HH:mm:ss");
			//cell.CellStyle = cellStyle; 
			#endregion


			#region 设置数值格式 - 保留2位小数
			//cell.SetCellValue(1.2);
			////numberformat with 2 digits after the decimal point - "1.20"
			//HSSFCellStyle cellStyle = hssfworkbook.CreateCellStyle() as HSSFCellStyle;
			//cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");
			//cell.CellStyle = cellStyle; 
			#endregion

			#region 设置货币格式
			//cell.SetCellValue(200676734.567777);
			//HSSFCellStyle cellStyle = hssfworkbook.CreateCellStyle() as HSSFCellStyle;
			//HSSFDataFormat format = hssfworkbook.CreateDataFormat() as HSSFDataFormat;
			////注意，这里还加入了千分位分隔符，所以是#,##，至于为什么这么写，你得去问微软，呵呵
			//cellStyle.DataFormat = format.GetFormat("¥#,##.00000");
			//cell.CellStyle = cellStyle; 
			#endregion

			#region 设置百分比
			//cell.SetCellValue(0.8);
			////numberformat with 2 digits after the decimal point - "1.20"
			//HSSFCellStyle cellStyle = hssfworkbook.CreateCellStyle() as HSSFCellStyle;
			//cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00%");
			//cell.CellStyle = cellStyle;  
			#endregion

			#region 设置数字转中文大小写
			//cell.SetCellValue(34654); //转换结果为：叁肆陆伍肆
			//HSSFDataFormat format = hssfworkbook.CreateDataFormat() as HSSFDataFormat;
			//HSSFCellStyle cellStyle = hssfworkbook.CreateCellStyle() as HSSFCellStyle;
			//cellStyle.DataFormat = format.GetFormat("[DbNum2][$-804]0");
			//cell.CellStyle = cellStyle;
			#endregion

			#region 科学计数法
			//cell.SetCellValue(346540000);
			//HSSFDataFormat format = hssfworkbook.CreateDataFormat() as HSSFDataFormat;
			//HSSFCellStyle cellStyle = hssfworkbook.CreateCellStyle() as HSSFCellStyle;
			//cellStyle.DataFormat = format.GetFormat("0.00E+00");
			//cell.CellStyle = cellStyle;
			#endregion


			#region 单元格合并（Region，合并单元格，其实就是设定一个区域。）
			/*
			 *	Region的参数		说明
				FirstRow		区域中第一个单元格的行号
				FirstColumn		区域中第一个单元格的列号
				LastRow			区域中最后一个单元格的行号
				LastColumn		区域中最后一个单元格的列号
			 */

			#region 行的合并
			//建立一张销售情况表，英文叫Sales Report
			//居中和字体样式，这里我们采用20号字体
			//cell.SetCellValue("Sales Report");
			//HSSFCellStyle style = hssfworkbook.CreateCellStyle() as HSSFCellStyle;

			////对齐方式
			//style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
			//style.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;

			//HSSFFont font = hssfworkbook.CreateFont() as HSSFFont;
			//font.FontHeight = 20 * 20; //设置字号
			//style.SetFont(font);
			////自动换行翻译成英文其实就是Wrap的意思，所以这里我们应该用WrapText属性，这是一个布尔属性
			//style.WrapText = true;
			////这是一个不太引人注意的选项，所以这里给张图出来，让大家知道是什么，缩进说白了就是文本前面的空白，我们同样可以用属性来设置，这个属性叫做Indention。
			//style.Indention = 3;
			////文本旋转
			////文本方向大家一定在Excel中设置过，上图中就是调整界面，主要参数是度数，那么我们如何在NPOI中设置呢？
			////这里的Rotation取值是从 - 90到90，而不是0 - 180度。
			//style.Rotation = 45;
			//cell.CellStyle = style;
			////要产生图中的效果，即把A1: F1这6个单元格合并，然后添加合并区域：
			//sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 6, 0, 5));
			#endregion

			#endregion



			#region 设置单元格边框
			/*
				边框相关属性			说明			范例				
				Border+方向			边框类型		BorderTop, BorderBottom,BorderLeft, BorderRight
				方向+BorderColor		边框颜色		TopBorderColor,BottomBorderColor, LeftBorderColor, RightBorderColor
			 */
			//HSSFCellStyle style = hssfworkbook.CreateCellStyle() as HSSFCellStyle;
			//style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
			//style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
			//style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
			//style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
			////以上代码将底部边框设置为绿色，要注意，不是直接把HSSFColor.GREEN赋给XXXXBorderColor属性，而是把index的值赋给它。
			//style.BottomBorderColor = HSSFColor.Green.Index;
			//cell.CellStyle = style;
			#endregion



			#region 设置单元格字体
			//cell.SetCellValue("Sales Report 微软雅黑");
			//HSSFFont font = hssfworkbook.CreateFont() as HSSFFont;
			//font.FontName = "微软雅黑"; //字体
			//font.IsBold = true;
			///* 说明：与字号有关的属性有两个，一个是FontHeight，一个是FontHeightInPoints。
			// * 区别在于，FontHeight的值是FontHeightInPoints的20倍，通常我们在Excel界面中看到的字号，
			// * 比如说12，对应的是FontHeightInPoints的值，而FontHeight要产生12号字体的大小，
			// * 值应该是240。所以通常建议你用FontHeightInPoint属性。
			// */
			////font.FontHeight = 12 * 12;
			//font.FontHeightInPoints = 18; //字号（建议使用该属性）
			//font.Color = HSSFColor.Red.Index; //字体颜色
			//font.Underline = NPOI.SS.UserModel.FontUnderlineType.Single; //下划线

			//#region 上标下标
			///*
			//	TypeOffset的值		说明
			//	HSSFFont.SS_SUPER	上标
			//	HSSFFont.SS_SUB		下标
			//	HSSFFont.SS_NONE	普通，默认值
			// */
			//font.TypeOffset = NPOI.SS.UserModel.FontSuperScript.Super;
			//#endregion

			//font.IsStrikeout = true; //删除线

			//HSSFCellStyle style = hssfworkbook.CreateCellStyle() as HSSFCellStyle;
			//style.SetFont(font);

			//cell.CellStyle = style;
			#endregion


			#region 设置单元格的背景和图案
			//cell.SetCellValue("Sales Report 微软雅黑");
			//HSSFCellStyle style = hssfworkbook.CreateCellStyle() as HSSFCellStyle;
			//style.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.White.Index;
			//style.FillPattern = NPOI.SS.UserModel.FillPattern.Squares; //图案样式
			//style.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Red.Index;
			//cell.CellStyle = style;
			#endregion


			#region 设置单元格的宽度和高度
			/*
			 * 这里你会发现一个有趣的现象，SetColumnWidth的第二个参数要乘以256，
			 * 这是怎么回事呢？其实，这个参数的单位是1/256个字符宽度，
			 * 也就是说，这里是把B列的宽度设置为了100个字符。
			 */
			//cell.SetCellValue("Sales Report 微软雅黑");
			//sheet.SetColumnWidth(0, 50 * 256); //设置宽度
			// var width= sheet.GetColumnWidth(0); //获取宽度


			///*
			// * 在Excel中，每一行的高度也是要求一致的，所以设置单元格的高度，
			// * 其实就是设置行的高度，所以相关的属性也应该在HSSFRow上，
			// * 它就是HSSFRow.Height和HeightInPoints，这两个属性的区别在于HeightInPoints的单位是点，
			// * 而Height的单位是1/20个点，所以Height的值永远是HeightInPoints的20倍。
			// */
			////row.Height = 200 * 20;
			//row.HeightInPoints = 200; //设置/获取高度

			////统一设置默认高度和宽度
			//sheet.DefaultColumnWidth = 100 * 256;
			//sheet.DefaultRowHeight = 30 * 20;
			#endregion

			#region 基本计算
			//HSSFSheet sheet1 = hssfworkbook.CreateSheet("Sheet1") as HSSFSheet;
			//HSSFRow row1 = sheet1.CreateRow(0) as HSSFRow;
			//var cel1 = row1.CreateCell(0);
			//var cel2 = row1.CreateCell(1);
			//var cel3 = row1.CreateCell(2);
			//cel1.SetCellFormula("1+2*3");
			//cel2.SetCellValue(5);
			///*
			// * NPOI也支持单元格引用类型的公式设置，如下图中的C1=A1*B1。
			// * 对应的公式设置代码为：cel3.SetCellFormula("A1*B1");
			// * 但要注意，在利用NPOI写程序时，行和列的计数都是从0开始计算的，
			// * 但在设置公式时又是按照Excel的单元格命名规则来的。
			// */
			//cel3.SetCellFormula("A1*B1");
			#endregion

			#region SUM函数（求和）
			//HSSFSheet sheet1 = hssfworkbook.CreateSheet("Sheet1") as HSSFSheet;
			//var row1 = sheet1.CreateRow(0);
			//var cel1 = row1.CreateCell(0);
			//var cel2 = row1.CreateCell(1);
			//var cel3 = row1.CreateCell(2);
			//var celSum1 = row1.CreateCell(3);
			//var celSum2 = row1.CreateCell(4);
			//var celSum3 = row1.CreateCell(5);

			//cel1.SetCellValue(1);
			//cel2.SetCellValue(2);
			//cel3.SetCellValue(3);
			///*
			// * 当然，把每一个单元格作为Sum函数的参数很容易理解，但如果要求和的单元格很多，那么公式就会很长，
			// * 既不方便阅读也不方便书写。所以Excel提供了另外一种多个单元格求和的写法：
			// * “Sum(A1:C1)”表示求从A1到C1所有单元格的和，相当于A1+B1+C1。
			// */
			//celSum2.SetCellFormula("sum(A1,C1)");

			///*
			// * 最后，还有一种求和的方法。就是先定义一个区域，如”range1”，
			// * 然后再设置Sum(range1)，此时将计算区域中所有单元格的和。
			// * 定义区域的代码为：
			// */
			//HSSFName range = hssfworkbook.CreateName() as HSSFName;
			//range.RefersToFormula = "Sheet1!$A1:$C1";
			//range.NameName = "range1";
			//celSum3.SetCellFormula("sum(range1)");
			#endregion

			#region 日期函数
			//HSSFSheet sheet1 = hssfworkbook.CreateSheet("Sheet1") as HSSFSheet;

			//var row1 = sheet1.CreateRow(0);
			//var row2 = sheet1.CreateRow(1);
			//row1.CreateCell(0).SetCellValue("姓名");
			//row1.CreateCell(1).SetCellValue("参加工作时间");
			//row1.CreateCell(2).SetCellValue("当前日期");
			//row1.CreateCell(3).SetCellValue("工作年限");

			//var cel1 = row2.CreateCell(0);
			//var cel2 = row2.CreateCell(1);
			//var cel3 = row2.CreateCell(2);
			//var cel4 = row2.CreateCell(3);

			//cel1.SetCellValue("aTao.Xiang");
			//cel2.SetCellValue(new DateTime(2004, 7, 1));
			//cel3.SetCellFormula("TODAY()");
			//cel4.SetCellFormula("CONCATENATE(DATEDIF(B2,TODAY(),\"y\"),\"年\",DATEDIF(B2,TODAY(),\"ym\"),\"个月\")");

			////在poi中日期是以double类型表示的，所以要格式化
			//HSSFCellStyle cellStyle = hssfworkbook.CreateCellStyle() as HSSFCellStyle;
			//HSSFDataFormat format = hssfworkbook.CreateDataFormat() as HSSFDataFormat;
			//cellStyle.DataFormat = format.GetFormat("yyyy-MM-dd");
			//cel2.CellStyle = cellStyle;
			//cel3.CellStyle = cellStyle;

			//sheet1.SetColumnWidth(1, 20 * 256); //设置宽度
			//sheet1.SetColumnWidth(2, 20 * 256); //设置宽度

			///*
			// * 下面对上例中用到的几个主要函数作一些说明：
			// * TODAY()：取得当前日期;
			// * DATEDIF(B2,TODAY(),"y")：取得B2单元格的日期与前日期以年为单位的时间间隔。
			// *	(“Y”:表示以年为单位,”m”表示以月为单位;”d”表示以天为单位);
			// *	CONCATENATE(str1,str2,...)：连接字符串。
			// */
			#endregion

			#region 画线
			//HSSFSheet sheet1 = hssfworkbook.CreateSheet("Sheet1") as HSSFSheet;
			//HSSFPatriarch patriarch = sheet1.CreateDrawingPatriarch() as HSSFPatriarch;
			//HSSFClientAnchor a1 = new HSSFClientAnchor(255, 125, 1023, 150, 0, 0, 2, 2);
			//HSSFSimpleShape line1 = patriarch.CreateSimpleShape(a1);

			//line1.ShapeType = HSSFSimpleShape.OBJECT_TYPE_LINE;
			//line1.LineStyle = HSSFShape.LINESTYLE_SOLID;
			////在NPOI中线的宽度12700表示1pt,所以这里是0.5pt粗的线条。
			//line1.LineWidth = 6350;
			#endregion

			#region 数据有效性
			//HSSFSheet sheet1 = hssfworkbook.CreateSheet("Sheet1") as HSSFSheet;

			//sheet1.CreateRow(0).CreateCell(0).SetCellValue("日期列");
			//CellRangeAddressList regions1 = new CellRangeAddressList(1, 65535, 0, 0);
			//DVConstraint constraint1 = DVConstraint.CreateDateConstraint(0, "1900-01-01", "2999-12-31", "yyyy-MM-dd");
			//HSSFDataValidation dataValidate1 = new HSSFDataValidation(regions1, constraint1);
			//dataValidate1.CreateErrorBox("error", "You must input a date.");
			//sheet1.AddValidationData(dataValidate1);
			#endregion

			#region 生成下拉列表
			//var sheet1 = hssfworkbook.CreateSheet("Sheet1");
			///*
			// *  先设置一个需要提供下拉的区域，关于CellRangeAddressList构造函数参数的说明请参见上一节：
			//	CellRangeAddressList regions = new CellRangeAddressList(0, 65535, 0, 0);
			//	然后将下拉项作为一个数组传给CreateExplicitListConstraint作为参数创建一个约束，根据要控制的区域和约束创建数据有效性就可以了。
			// */
			//CellRangeAddressList regions = new CellRangeAddressList(0, 65535, 0, 0);
			//DVConstraint constraint = DVConstraint.CreateExplicitListConstraint(new string[] { "itemA", "itemB", "itemC" });
			//HSSFDataValidation dataValidate = new HSSFDataValidation(regions, constraint);
			//sheet1.AddValidationData(dataValidate);
			#endregion

			#region 生成九九乘法表
			//var sheet1 = hssfworkbook.CreateSheet("Sheet1");
			//HSSFRow row;
			//HSSFCell cell;
			//for (int rowIndex = 0; rowIndex < 9; rowIndex++)
			//{
			//	row = sheet1.CreateRow(rowIndex) as HSSFRow;
			//	for (int colIndex = 0; colIndex <= rowIndex; colIndex++)
			//	{
			//		cell = row.CreateCell(colIndex) as HSSFCell;
			//		cell.SetCellValue(String.Format("{0}*{1}={2}", rowIndex + 1, colIndex + 1, (rowIndex + 1) * (colIndex + 1)));
			//	}
			//}
			#endregion

			#region 生成一张工资单
			////写标题文本
			//HSSFSheet sheet1 = hssfworkbook.CreateSheet("Sheet1") as HSSFSheet;

			//HSSFCell cellTitle = sheet1.CreateRow(0).CreateCell(0) as HSSFCell;
			//cellTitle.SetCellValue("XXX公司2009年10月工资单");

			////设置标题行样式
			//HSSFCellStyle style = hssfworkbook.CreateCellStyle() as HSSFCellStyle;
			//style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
			//style.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
			//HSSFFont font = hssfworkbook.CreateFont() as HSSFFont;
			//font.FontHeight = 20 * 20;
			//style.SetFont(font);

			//cellTitle.CellStyle = style;

			////合并标题行
			//sheet1.AddMergedRegion(new CellRangeAddress(0, 1, 0, 6));

			////其中用到了我们前面讲的设置单元格样式和合并单元格等内容。接下来我们循环创建公司每个员工的工资单：

			//DataTable dt = GetData();
			//var row = sheet1.CreateRow(2);
			//var cell = row.CreateCell(2);
			//HSSFCellStyle celStyle = getCellStyle(hssfworkbook);

			//HSSFPatriarch patriarch = sheet1.CreateDrawingPatriarch() as HSSFPatriarch;
			//HSSFClientAnchor anchor;
			//HSSFSimpleShape line;
			//int rowIndex;
			//for (int i = 0; i < dt.Rows.Count; i++)
			//{
			//	//表头数据
			//	rowIndex = 3 * (i + 1);
			//	row = sheet1.CreateRow(rowIndex);

			//	cell = row.CreateCell(0);
			//	cell.SetCellValue("姓名");
			//	cell.CellStyle = celStyle;

			//	cell = row.CreateCell(1);
			//	cell.SetCellValue("基本工资");
			//	cell.CellStyle = celStyle;

			//	cell = row.CreateCell(2);
			//	cell.SetCellValue("住房公积金");
			//	cell.CellStyle = celStyle;

			//	cell = row.CreateCell(3);
			//	cell.SetCellValue("绩效奖金");
			//	cell.CellStyle = celStyle;

			//	cell = row.CreateCell(4);
			//	cell.SetCellValue("社保扣款");
			//	cell.CellStyle = celStyle;

			//	cell = row.CreateCell(5);
			//	cell.SetCellValue("代扣个税");
			//	cell.CellStyle = celStyle;

			//	cell = row.CreateCell(6);
			//	cell.SetCellValue("实发工资");
			//	cell.CellStyle = celStyle;


			//	DataRow dr = dt.Rows[i];
			//	//设置值和计算公式
			//	row = sheet1.CreateRow(rowIndex + 1);
			//	cell = row.CreateCell(0);
			//	cell.SetCellValue(dr["FName"].ToString());
			//	cell.CellStyle = celStyle;

			//	cell = row.CreateCell(1);
			//	cell.SetCellValue((double)dr["FBasicSalary"]);
			//	cell.CellStyle = celStyle;

			//	cell = row.CreateCell(2);
			//	cell.SetCellValue((double)dr["FAccumulationFund"]);
			//	cell.CellStyle = celStyle;

			//	cell = row.CreateCell(3);
			//	cell.SetCellValue((double)dr["FBonus"]);
			//	cell.CellStyle = celStyle;

			//	cell = row.CreateCell(4);
			//	cell.SetCellFormula(String.Format("$B{0}*0.08", rowIndex + 2));
			//	cell.CellStyle = celStyle;

			//	cell = row.CreateCell(5);
			//	cell.SetCellFormula(String.Format("SUM($B{0}:$D{0})*0.1", rowIndex + 2));
			//	cell.CellStyle = celStyle;

			//	cell = row.CreateCell(6);
			//	cell.SetCellFormula(String.Format("SUM($B{0}:$D{0})-SUM($E{0}:$F{0})", rowIndex + 2));
			//	cell.CellStyle = celStyle;


			//	//绘制分隔线
			//	sheet1.AddMergedRegion(new CellRangeAddress(rowIndex + 2,rowIndex + 2, 0,  6));
			//	anchor = new HSSFClientAnchor(0, 125, 1023, 125, 0, rowIndex + 2, 6, rowIndex + 2);
			//	line = patriarch.CreateSimpleShape(anchor);
			//	line.ShapeType = HSSFSimpleShape.OBJECT_TYPE_LINE;
			//	line.LineStyle = NPOI.SS.UserModel.LineStyle.DashGel;

			//}

			#endregion

			WriteXlsToFile(hssfworkbook, @"test.xls");
		}

		static DataTable GetData()
		{
			DataTable dt = new DataTable();
			dt.Columns.Add("FName", typeof(System.String));
			dt.Columns.Add("FBasicSalary", typeof(System.Double));
			dt.Columns.Add("FAccumulationFund", typeof(System.Double));
			dt.Columns.Add("FBonus", typeof(System.Double));

			dt.Rows.Add("令狐冲", 6000, 1000, 2000);
			dt.Rows.Add("任盈盈", 7000, 1000, 2500);
			dt.Rows.Add("林平之", 5000, 1000, 1500);
			dt.Rows.Add("岳灵珊", 4000, 1000, 900);
			dt.Rows.Add("任我行", 4000, 1000, 800);
			dt.Rows.Add("风清扬", 9000, 5000, 3000);

			return dt;
		}

		static HSSFCellStyle getCellStyle(HSSFWorkbook hssfworkbook)
		{
			HSSFCellStyle cellStyle = hssfworkbook.CreateCellStyle() as HSSFCellStyle;
			cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
			cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
			cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
			cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
			return cellStyle;
		}

		/// <summary>
		/// 写入文件
		/// </summary>
		/// <param name="hssfworkbook"></param>
		/// <param name="path"></param>
		private void WriteXlsToFile(HSSFWorkbook hssfworkbook, string path)
		{
			using (FileStream file = new FileStream(path, FileMode.Create))
			{
				hssfworkbook.Write(file);
			}
		}
	}
}
