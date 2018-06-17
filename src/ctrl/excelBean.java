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
 * ����jacob����MSexcel�� �����word�������ƣ������н����������˼�
 */

public class excelBean extends bean
{
	// ָ����幤����
	private Dispatch xls;

	// excel���г������
	private ActiveXComponent excel;

	// ����������
	private Dispatch workbooks = null;

	// �������б�
	private Dispatch sheet = null;

	// ���캯������ʼ��excelӦ��
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
	 * ��һ���Ѵ��ڵ�excel�ĵ�
	 * 
	 * @param xlsPath xls�ĵ�·��
	 */
	public void openDocument(String xlsPath)
	{
		xls = Dispatch.invoke(workbooks, "Open", Dispatch.Method, new Object[]
		{ xlsPath, new Variant(false), new Variant(true) }, // �Ƿ���ֻ����ʽ��
				new int[1]).toDispatch();
		// ��ȡ��ǰ�Ĺ�����
		sheet = Dispatch.get(xls, "ActiveSheet").toDispatch();
	}

	/**
	 * ֱ�ӱ�����Ŀ��·��
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
	 * ����µĹ�����(sheet)������Ӻ�ΪĬ��Ϊ��ǰ����Ĺ�����
	 */
	public Dispatch addSheet()
	{
		return Dispatch.get(Dispatch.get(xls, "sheets").toDispatch(), "add").toDispatch();
	}

	/**
	 * �޸ĵ�ǰ�����������
	 * 
	 * @param newName
	 */
	public void modifyCurrentSheetName(String newName)
	{
		Dispatch.put(sheet, "name", newName);
	}

	/**
	 * �õ���ǰ�����������
	 * 
	 * @return
	 */
	public String getCurrentSheetName()
	{
		return Dispatch.get(sheet, "name").toString();
	}

	/**
	 * �õ�������������
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
	 * ͨ�����������ֵõ�������
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
	 * ͨ�������������õ�������(��һ��������indexΪ1)
	 * 
	 * @param index
	 * @return sheet����
	 */
	public Dispatch getSheetByIndex(Integer index)
	{
		return Dispatch.invoke(sheet, "Item", Dispatch.Get, new Object[]
		{ index }, new int[1]).toDispatch();
	}

	/**
	 * �õ������������
	 * 
	 * @return
	 */
	public int getSheetCount()
	{
		int count = Dispatch.get(workbooks, "count").toInt();
		return count;
	}

	/**
	 * д������
	 * 
	 * @param cells �洢�Ÿ�����Ԫ����Ϣ
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
	 * ������󣬹ر��ļ�
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
	 * д��ֵ
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
	 * ��ȡֵ
	 * 
	 * @param position ���� c1
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
	 * ��ȡ�������
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
	 * ��ȡ�������
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
	 * ��ȡλ��
	 * 
	 * @param rnum �������
	 * @param cnum �������
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
	 * �ڲ��࣬��ʾexcel�ĵ�Ԫ������
	 * 
	 * @author Wenzhou
	 *
	 */
	private class ExcelCell
	{

		private int row; // �к�
		private int col; // �к�
		private String value; // ��Ӧֵ

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
	 * �ļ���ʽת��������Ҳ�������Ϊһ��
	 * ��ת�������ͣ����https://msdn.microsoft.com/zh-cn/VBA/Excel-VBA/articles/xlfileformat-enumeration-excel
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
		            new Variant(0), //PDF��ʽ=0  
		            writePath,  
		            new Variant(0)  //0=��׼ (���ɵ�PDFͼƬ�����ģ��) 1=��С�ļ� (���ɵ�PDFͼƬ����һ����Ϳ)  
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
