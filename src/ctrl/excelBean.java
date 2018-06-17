package ctrl;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 利用jacob操作MSexcel类 框架与word基本相似，本类中将操作进行了简化
 */

public class excelBean extends bean
{
	// 指向具体工作簿
	private Dispatch xls;

	// excel运行程序对象
	private ActiveXComponent excel;

	// 工作簿集合
	private Dispatch workbooks = null;

	// 工作簿中表单
	private Dispatch sheet = null;

	// 构造函数，初始化excel应用
	public excelBean() throws Exception
	{
		ComThread.InitSTA();

		if (excel == null)
		{
			excel = new ActiveXComponent("Excel.Application");
			excel.setProperty("Visible", new Variant(false));
			excel.setProperty("AutomationSecurity", new Variant(3));
		}
		if (workbooks == null)
		{
			workbooks = excel.getProperty("Workbooks").toDispatch();
		}
	}

	/**
	 * 打开一个已存在的excel文档
	 * 
	 * @param xlsPath xls文档路径
	 */
	public void openDocument(String xlsPath)
	{
		xls = Dispatch.invoke(workbooks, "Open", Dispatch.Method, new Object[]
		{ xlsPath, new Variant(false), new Variant(true) }, // 是否以只读方式打开
				new int[1]).toDispatch();
		// 获取当前的工作表
		sheet = Dispatch.get(xls, "ActiveSheet").toDispatch();
	}

	/**
	 * 直接保存至目标路径
	 * 
	 * @param savePath
	 */
	public void save(String savePath)
	{
		// Dispatch.call(Dispatch.call(excel, "WordBasic").getDispatch(),
		// "FileSaveAs", savePath);
		Dispatch.call(xls, "SaveAs", savePath);
	}

	/**
	 * 添加新的工作表(sheet)，（添加后为默认为当前激活的工作表）
	 */
	public Dispatch addSheet()
	{
		return Dispatch.get(Dispatch.get(xls, "sheets").toDispatch(), "add").toDispatch();
	}

	/**
	 * 修改当前工作表的名字
	 * 
	 * @param newName
	 */
	public void modifyCurrentSheetName(String newName)
	{
		Dispatch.put(sheet, "name", newName);
	}

	/**
	 * 得到当前工作表的名字
	 * 
	 * @return
	 */
	public String getCurrentSheetName()
	{
		return Dispatch.get(sheet, "name").toString();
	}

	/**
	 * 得到工作薄的名字
	 * 
	 * @return
	 */
	public String getWorkbookName()
	{
		if (xls == null)
			return null;
		return Dispatch.get(xls, "name").toString();
	}

	/**
	 * 通过工作表名字得到工作表
	 * 
	 * @param name sheetName
	 * @return
	 */
	public Dispatch getSheetByName(String name)
	{
		return Dispatch.invoke(sheet, "Item", Dispatch.Get, new Object[]
		{ name }, new int[1]).toDispatch();
	}

	/**
	 * 通过工作表索引得到工作表(第一个工作簿index为1)
	 * 
	 * @param index
	 * @return sheet对象
	 */
	public Dispatch getSheetByIndex(Integer index)
	{
		return Dispatch.invoke(sheet, "Item", Dispatch.Get, new Object[]
		{ index }, new int[1]).toDispatch();
	}

	/**
	 * 得到工作表的总数
	 * 
	 * @return
	 */
	public int getSheetCount()
	{
		int count = Dispatch.get(workbooks, "count").toInt();
		return count;
	}

	/**
	 * 写入数据
	 * 
	 * @param cells 存储着各个单元格信息
	 */
	public void putData(List<ExcelCell> cells) throws Exception
	{
		if (cells == null || cells.size() == 0)
		{
			return;
		}

		String position = null;
		int row = 0;
		int col = 0;
		int max = 26 * 26 - 1;
		for (ExcelCell c : cells)
		{
			row = c.getRow();
			col = c.getCol();
			if (row < 0 || col < 0 || col > max || c.getValue() == null || c.getValue().trim().equals(""))
			{
				continue;
			}
			setValue(position, c.getValue());
		}
	}

	/**
	 * 保存过后，关闭文件
	 */
	public void close()
	{
		// Dispatch.call(xls, "Save");
		Dispatch.call(xls, "Close", new Variant(false));
		excel.invoke("Quit", new Variant[]
		{});
		ComThread.Release();
	}

	/**
	 * 写入值
	 * 
	 * @param position
	 * @param value
	 */
	protected void setValue(String position, String value)
	{
		Dispatch cell = Dispatch.invoke(sheet, "Range", Dispatch.Get, new Object[]
		{ position }, new int[1]).toDispatch();
		Dispatch.put(cell, "Value", value);
	}

	/**
	 * 读取值
	 * 
	 * @param position 例如 c1
	 * @return
	 */
	protected String getValue(String position)
	{
		// System.out.println(position);
		Dispatch cell = Dispatch.invoke(sheet, "Range", Dispatch.Get, new Object[]
		{ position }, new int[1]).toDispatch();
		String value = Dispatch.get(cell, "Value").toString();
		return value;
	}

	/**
	 * 获取最大行数
	 * 
	 * @return
	 */
	private int getRowCount()
	{
		Dispatch UsedRange = Dispatch.get(sheet, "UsedRange").toDispatch();
		Dispatch rows = Dispatch.get(UsedRange, "Rows").toDispatch();
		int num = Dispatch.get(rows, "count").getInt();
		return num;
	}

	/**
	 * 获取最大列数
	 * 
	 * @return
	 */
	private int getColumnCount()
	{
		Dispatch UsedRange = Dispatch.get(sheet, "UsedRange").toDispatch();
		Dispatch Columns = Dispatch.get(UsedRange, "Columns").toDispatch();
		int num = Dispatch.get(Columns, "count").getInt();
		return num;
	}

	/**
	 * 获取位置
	 * 
	 * @param rnum 最大行数
	 * @param cnum 最大列数
	 */
	private String getCellPosition(int rnum, int cnum)
	{
		String cposition = "";
		if (cnum > 26)
		{
			int multiple = (cnum) / 26;
			int remainder = (cnum) % 26;
			char mchar = (char) (multiple + 64);
			char rchar = (char) (remainder + 64);
			cposition = mchar + "" + rchar;
		} else
		{
			cposition = (char) (cnum + 64) + "";
		}
		cposition += rnum;
		return cposition;
	}

	/**
	 * 内部类，表示excel的单元格内容
	 * 
	 * @author Wenzhou
	 *
	 */
	private class ExcelCell
	{

		private int row; // 行号
		private int col; // 列号
		private String value; // 对应值

		public ExcelCell()
		{

		}

		public ExcelCell(int row, int col, String value)
		{
			this.row = row;
			this.col = col;
			this.value = value;
		}

		public int getRow()
		{
			return row;
		}

		public void setRow(int row)
		{
			this.row = row;
		}

		public int getCol()
		{
			return col;
		}

		public void setCol(int col)
		{
			this.col = col;
		}

		public String getValue()
		{
			return value;
		}

		public void setValue(String value)
		{
			this.value = value;
		}
	}

	public ArrayList<String[]> getContext()
	{
		ArrayList<String[]> dataList = new ArrayList<>();
		int col = getColumnCount();
		int row = getRowCount();
		for (int i = 0; i < row; i++)
		{
			String[] items = new String[col];
			for (int j = 0; j < col; j++)
			{
				items[j] = getValue(getCellPosition(i + 1, j + 1));
				System.out.println(items[j]);
			}
			dataList.add(items);
		}
		return dataList;
	}
	
	@Override
	/**
	 * 文件格式转换函数，也就是另存为一下
	 * 可转化的类型，详见https://msdn.microsoft.com/zh-cn/VBA/Excel-VBA/articles/xlfileformat-enumeration-excel
	 */
	public void conversion(String readPath, String curFormat, String writePath, String tarFormat)
	{
		String[] formatName = {"xls", "xlsx", "html", "csv", "txt", "xml"};
		int[] formatId = {56, 51, 44, 6, -4158, 46};
		Map<String, Integer> formatTo = new HashMap<String, Integer>(); 
		
		for(int i=0; i<formatName.length; i++)
		{
			formatTo.put(formatName[i], formatId[i]);
		}
		this.openDocument(readPath);
//		System.out.println(this.getWorkbookName());
		if(tarFormat.equals("pdf")==false)
		{
			Dispatch.invoke(xls,"SaveAs",Dispatch.Method,new Object[]{
	            writePath, new Variant(formatTo.get(tarFormat)),},new int[1]);
//			Dispatch.call(xls, "SaveAs", savePath);

		}
		else
		{
			Dispatch.invoke(xls,"ExportAsFixedFormat",Dispatch.Method,new Object[]{  
		            new Variant(0), //PDF格式=0  
		            writePath,  
		            new Variant(0)  //0=标准 (生成的PDF图片不会变模糊) 1=最小文件 (生成的PDF图片糊的一塌糊涂)  
		        },new int[1]);
//			Dispatch.call((Dispatch)workbook, "SaveAs", nameFile ,new Variant(2));
		}
	}
/*
	public static void main(String[] args)
	{
		excelBean eb = null;
		try
		{
			eb = new excelBean();
			eb.conversion("D:\\Documents\\JAVA\\Transfer\\src\\testdata\\yancheng.xls", "xls", "D:\\Documents\\JAVA\\Transfer\\src\\testdata\\yancheng.pdf", "pdf");
			eb.close();
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
			eb.close();
		}

//		eb.openDocument("D:\\Documents\\JAVA\\Transfer\\src\\testdata\\yancheng.xlsx");
		// eb.setValue("C2", "value");
//		 eb.save("C:\\Users\\pc\\JAVA\\Transfer\\src\\testdata\\yancheng2.xlsx");
	}
	*/
}
