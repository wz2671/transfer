package ctrl;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class pptBean extends bean
{
	public static final int ppSaveAsJPG = 17;

	private static final String ADD_CHART = "AddChart";

	private ActiveXComponent ppt;
	private ActiveXComponent presentation;

	/**
	 * ����һ���µ�PPT
	 * 
	 * @param isVisble
	 */
	public pptBean(boolean isVisble)
	{
		if (null == ppt)
		{
			ppt = new ActiveXComponent("PowerPoint.Application");
			ppt.setProperty("Visible", new Variant(isVisble));
			ActiveXComponent presentations = ppt.getPropertyAsComponent("Presentations");
			presentation = presentations.invokeGetComponent("Add", new Variant(1));
		}
	}

	public pptBean(String filePath, boolean isVisble) throws Exception
	{
		if (null == filePath || "".equals(filePath))
		{
			throw new Exception("�ļ�·��Ϊ��!");
		}
		File file = new File(filePath);
		if (!file.exists())
		{
			throw new Exception("�ļ�������!");
		}
		ppt = new ActiveXComponent("PowerPoint.Application");
		try
		{
			ppt.setProperty("Visible", new Variant(isVisble));
		}
		catch(Exception e)
		{
			System.out.println("������ʾ");
		}
		// ��һ�����е� Presentation ����
		ActiveXComponent presentations = ppt.getPropertyAsComponent("Presentations");
		presentation = presentations.invokeGetComponent("Open", new Variant(filePath), new Variant(true));
	}

	/**
	 * ��pptת��ΪͼƬ
	 * 
	 * @param pptfile
	 * @param saveToFolder
	 */
	public void PPTToJPG(String pptfile, String saveToFolder)
	{
		try
		{
			saveAs(presentation, saveToFolder, ppSaveAsJPG);
			if (presentation != null)
			{
				Dispatch.call(presentation, "Close");
			}
		} catch (Exception e)
		{
			ComThread.Release();
		} finally
		{
			if (presentation != null)
			{
				Dispatch.call(presentation, "Close");
			}
			ppt.invoke("Quit", new Variant[]
			{});
			ComThread.Release();
		}
	}

	/**
	 * ����ppt
	 * 
	 * @param pptFile
	 * @date 2009-7-4
	 */
	public void PPTShow(String pptFile)
	{
		// powerpoint�õ�չʾ���ö���
		ActiveXComponent setting = presentation.getPropertyAsComponent("SlideShowSettings");
		// ���øö����run����ʵ��ȫ������
		setting.invoke("Run");
		// �ͷſ����߳�
		ComThread.Release();
	}

	/**
	 * ppt���Ϊ
	 * 
	 * @param presentation
	 * @param saveTo
	 * @param ppSaveAsFileType
	 */
	public void saveAs(Dispatch presentation, String saveTo, int ppSaveAsFileType)
	{
		Dispatch.call(presentation, "SaveAs", saveTo, new Variant(ppSaveAsFileType));
	}

	/**
	 * �ر�PPT���ͷ��߳�
	 * 
	 * @throws Exception
	 */
	public void closePpt() throws Exception
	{
		if (null != presentation)
		{
			Dispatch.call(presentation, "Close");
		}
		ppt.invoke("Quit", new Variant[]
		{});
		ComThread.Release();
	}

	/**
	 * ����PPT
	 * 
	 * @throws Exception
	 */
	public void runPpt() throws Exception
	{
		ActiveXComponent setting = presentation.getPropertyAsComponent("SlideShowSettings");
		setting.invoke("Run");
	}

	/**
	 * �����Ƿ�ɼ�
	 * 
	 * @param visble
	 * @param obj
	 */
	private void setIsVisble(Dispatch obj, boolean visble) throws Exception
	{
		Dispatch.put(obj, "Visible", new Variant(visble));
	}

	/**
	 * �ڻõ�Ƭ����������µĻõ�Ƭ
	 * 
	 * @param slides
	 * @param pptPage
	 *            �õ�Ƭ���
	 * @param type
	 *            4:����+��� 2:����+�ı� 3:����+���ҶԱ��ı� 5:����+���ı���ͼ�� 6:����+��ͼ�����ı�
	 *            7:����+SmartArtͼ�� 8:����+ͼ��
	 * @return
	 * @throws Exception
	 */
	private Variant addPptPage(ActiveXComponent slides, int pptPage, int type) throws Exception
	{
		return slides.invoke("Add", new Variant(pptPage), new Variant(type));
	}

	/**
	 * 
	 * @param pageShapes
	 *            ҳ���SHAPES�Ķ���
	 * @param chartType
	 *            ͼ������
	 * @param leftDistance
	 *            ������߿�ľ���
	 * @param topDistance
	 *            �����ϱ߿�ľ���
	 * @param width
	 *            ͼ��Ŀ��
	 * @param height
	 *            ͼ��ĸ߶�
	 * @return
	 * @throws Exception
	 */
	public Dispatch addChart(Dispatch pageShapes, int chartType, int leftDistance, int topDistance, int width,
			int height) throws Exception
	{
		Variant chart = Dispatch.invoke(pageShapes, ADD_CHART, 1, new Object[]
		{ new Integer(chartType), // ͼ������
				new Integer(leftDistance), // ������߿�ľ���
				new Integer(topDistance), // �����ϱ߿�ľ���
				new Integer(width), // ͼ��Ŀ��
				new Integer(height),// ͼ��ĸ߶�
		}, new int[1]);// ��������

		return chart.toDispatch();
	}

	/**
	 * ��ȡ�ڼ����õ�Ƭ
	 * 
	 * @param index
	 *            ��ţ���1��ʼ
	 * @return
	 * @throws Exception
	 */
	public Dispatch getPptPage(int pageIndex) throws Exception
	{
		// ��ȡ�õ�Ƭ����
		ActiveXComponent slides = presentation.getPropertyAsComponent("Slides");
		// ��õڼ���PPT
		Dispatch pptPage = Dispatch.call(slides, "Item", new Object[]
		{ new Variant(pageIndex) }).toDispatch();
		Dispatch.call(pptPage, "Select");
		return pptPage;
	}

	/**
	 * ��ӱ��
	 * 
	 * @param pageShapes
	 *            ҳ�����
	 * @param rows
	 *            ����
	 * @param columns
	 *            ����
	 * @param leftDistance
	 *            ������߾���
	 * @param topDistance
	 *            ���붥������
	 * @param width
	 *            ���
	 * @param height
	 *            �߶�
	 * @return
	 * @throws Exception
	 */
	public Dispatch addTable(Dispatch pageShapes, long rows, long columns, int leftDistance, int topDistance, int width,
			int height) throws Exception
	{
		return Dispatch.invoke(pageShapes, "AddTable", 1, new Object[]
		{ new Long(rows), new Long(columns), new Integer(leftDistance), new Integer(topDistance), new Integer(width),
				new Integer(height) }, new int[1]).toDispatch();
	}

	/**
	 * ��Selection�������޸�TEXT�����ֵ
	 * 
	 * @param selectionObj
	 * @param value
	 * @throws Exception
	 */
	public void addTextValue(Dispatch selectionObj, String value) throws Exception
	{
		Dispatch shapeRange = (Dispatch) Dispatch.get(selectionObj, "ShapeRange").getDispatch();
		Dispatch textFrame = (Dispatch) Dispatch.get(shapeRange, "TextFrame").getDispatch();
		Dispatch textRange = (Dispatch) Dispatch.get(textFrame, "TextRange").getDispatch();
		Dispatch.call(textRange, "Select");
		Dispatch.put(textRange, "Text", value);
	}

	/**
	 * ��������ӵ��ƶ��ĵ�Ԫ����
	 * 
	 * @param cell
	 *            ��Ԫ�����
	 * @param value
	 *            ��Ҫ��ӵ�����
	 * @throws Exception
	 */
	public void addCellValue(Dispatch cell, Object value) throws Exception
	{
		Dispatch cellShape = Dispatch.get(cell, "Shape").toDispatch();
		Dispatch cellFrame = Dispatch.get(cellShape, "TextFrame").toDispatch();
		Dispatch cellRange = Dispatch.get(cellFrame, "TextRange").toDispatch();
		Dispatch.put(cellRange, "Text", value);
	}

	/**
	 * �ϲ���Ԫ��,�ϲ�֮��ԭ��������Ԫ������ݽ��ŵ�һ����Ԫ������ �����ʼ��Ԫ��ͽ�����Ԫ֮��缸����Ԫ�񣬽���һ�𱻺ϲ�
	 * 
	 * @param cell
	 *            ��ʼ��Ԫ��
	 * @param cell2
	 *            ������Ԫ��
	 * @return
	 * @throws Exception
	 */
	public void mergeCell(Dispatch cell, Dispatch cell2) throws Exception
	{
		Dispatch.invoke(cell, "Merge", 1, new Object[]
		{ cell2 }, new int[1]);
	}

	/**
	 * ��ȡ�����ƶ���Ԫ��
	 * 
	 * @param tableObj
	 *            ������
	 * @param rowNum
	 *            �ڼ��У���1��ʼ
	 * @param columnRum
	 *            �ڼ��У���1��ʼ
	 * @return
	 * @throws Exception
	 */
	public Dispatch getCellOfTable(Dispatch tableObj, int rowNum, int columnRum) throws Exception
	{
		return Dispatch.invoke(tableObj, "Cell", Dispatch.Method, new Object[]
		{ new Long(rowNum), new Long(columnRum) }, new int[1]).toDispatch();
	}

	/**
	 * ���õ�Ԫ�񱳾���ɫ
	 * 
	 * @param cellObj
	 * @param colorIndex
	 * @throws Exception
	 */
	public void setCellBackColor(Dispatch cellObj, int colorIndex) throws Exception
	{
		Dispatch cellShape = Dispatch.get(cellObj, "Shape").toDispatch();
		Dispatch fillObj = Dispatch.get(cellShape, "Fill").toDispatch();
		Dispatch backColor = Dispatch.get(fillObj, "ForeColor").toDispatch();
		Dispatch.put(backColor, "ObjectThemeColor", colorIndex);
		Dispatch.put(fillObj, "ForeColor", backColor);
	}

	/**
	 * �޸ı�����ʽ��Ĭ����ʽΪ:{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}
	 * 
	 * @param tableObj
	 *            ������
	 * @param styleId
	 *            ��ʽID
	 * @throws Exception
	 */
	public void editTableSyle(Dispatch tableObj, String styleId) throws Exception
	{
		if (null == tableObj)
		{
			throw new Exception("��Ч�ı�����!");
		}
		if (null == styleId || "".equals(styleId))
		{
			throw new Exception("��Ч����ʽID!");
		}
		Dispatch.invoke(tableObj, "ApplyStyle", Dispatch.Method, new Object[]
		{ styleId }, new int[1]);
	}

	/**
	 * ��TABLE�����������
	 * 
	 * @param tableObj
	 * @param beforeColumn
	 * @throws Exception
	 */
	public void addTableColumn(Dispatch tableObj, int beforeColumn) throws Exception
	{
		Dispatch columns = Dispatch.get(tableObj, "Columns").getDispatch();
		int count = Dispatch.get(columns, "Count").getInt();
		if (beforeColumn > count || beforeColumn < 1)
		{
			throw new Exception("��Ч��������!");
		}
		Dispatch.invoke(columns, "Add", Dispatch.Method, new Object[]
		{ beforeColumn }, new int[1]);
	}

	/**
	 * ��TABLE�����������
	 * 
	 * @param tableObj
	 * @param beforeColumn
	 * @throws Exception
	 */
	public void addTableRow(Dispatch tableObj, int beforeRow) throws Exception
	{
		Dispatch rows = Dispatch.get(tableObj, "Rows").getDispatch();
		int count = Dispatch.get(rows, "Count").getInt();
		if (beforeRow > count || beforeRow < 1)
		{
			throw new Exception("��Ч��������!");
		}
		Dispatch.invoke(rows, "Add", Dispatch.Method, new Object[]
		{ beforeRow }, new int[1]);
	}

	/**
	 * �޸ĵ���Щ�е�ͼ������
	 * 
	 * @param chartObj
	 *            ͼ�����
	 * @param seriIndex
	 *            ϵ����������1��ʼ
	 * @param chartType
	 *            ͼ������
	 * @throws Exception
	 */
	public void updateSeriChartType(Dispatch chartObj, int seriIndex, int chartType) throws Exception
	{
		Dispatch Seri1 = Dispatch.call(chartObj, "SeriesCollection", new Variant(seriIndex)).toDispatch();
		Dispatch.put(Seri1, "ChartType", chartType);
	}

	/**
	 * �����Ƿ���ʾͼ������ݱ��,������һ�����ʱĬ��ʱ����ʾ��
	 * 
	 * @param chartObj
	 *            ������
	 * @param bValue
	 *            ����ֵ��tureΪ��ʾ��falseΪ����ʾ
	 * @throws Exception
	 */
	public void setIsDispDataTable(Dispatch chartObj, boolean bValue) throws Exception
	{
		if (null == chartObj)
		{
			throw new Exception("��Ч��ͼ�����!");
		}
		Dispatch.put(chartObj, "HasDataTable", bValue);
	}

	/**
	 * ��ȡ������ʽID
	 * 
	 * @param tableObj
	 * @return
	 * @throws Exception
	 */
	public String getTableStyleId(Dispatch tableObj) throws Exception
	{
		Dispatch tableStyle = Dispatch.get(tableObj, "Style").toDispatch();
		return Dispatch.get(tableStyle, "Id").toString();
	}

	/**
	 * ����ͼ�����Ƿ���ʾ���ݱ��
	 * 
	 * @param chartObj
	 * @param value
	 * @throws Exception
	 */
	public void setHasDataTable(Dispatch chartObj, boolean value) throws Exception
	{
		Dispatch.put(chartObj, "HasDataTable", value);
	}

	public void getGeneragePpt() throws Exception
	{
		// ����һ���µ�ppt ����
		Dispatch windows = presentation.getProperty("Windows").toDispatch();
		Dispatch window = Dispatch.call(windows, "Item", new Variant(1)).toDispatch();
		Dispatch selection = Dispatch.get(window, "Selection").toDispatch();
		// ��ȡ�õ�Ƭ����
		ActiveXComponent slides = presentation.getPropertyAsComponent("Slides");

		// ��ӵ�һ�Żõ�Ƭ�� ����+������
		addPptPage(slides, 1, 1);

		Dispatch slideRange = (Dispatch) Dispatch.get(selection, "SlideRange").getDispatch();
		Dispatch shapes = (Dispatch) Dispatch.get(slideRange, "Shapes").getDispatch();
		// ��ȡ�õ�Ƭ�еĵ�һ��Ԫ��
		Dispatch shape1 = Dispatch.call(shapes, "Item", new Variant(1)).toDispatch();
		// ��ȡ�õ�Ƭ�еĵڶ���Ԫ��
		Dispatch shape2 = Dispatch.call(shapes, "Item", new Variant(2)).toDispatch();
		// ѡ�е�һ��Ԫ��
		Dispatch.call(shape1, "Select");
		// ���ֵ
		addTextValue(selection, "����������");
		// ����PPTһҳ�еĵڶ���shape
		Dispatch.call(shape2, "Select");
		addTextValue(selection, "���Ը�����");
		// ��ӵڶ��Żõ�Ƭ(����+ͼ��)
		Variant v = addPptPage(slides, 2, 8);
		// ��ȡ�ڶ���PPT����
		Dispatch pptTwo = v.getDispatch();
		// ���ǰPPT����
		Dispatch.call(pptTwo, "Select");
		// ��ȡPPT�е�shapes
		shapes = Dispatch.get(pptTwo, "Shapes").toDispatch();
		Dispatch shapeText = Dispatch.call(shapes, "Item", new Variant(1)).toDispatch();
		// ��������
		Dispatch.call(shapeText, "Select");
		addTextValue(selection, "����ͼ�����");
		// ���ͼ��

		Dispatch chartDisp = addChart(shapes, 2, 10, 130, 700, 200);
		Dispatch chartObj = Dispatch.get(chartDisp, "Chart").getDispatch();

		setHasDataTable(chartObj, true);

		Dispatch chartData = Dispatch.get(chartObj, "ChartData").toDispatch();

		// Variant chartStyleV=Dispatch.get(chartObj, "ChartFont");

		Dispatch workBook = Dispatch.get(chartData, "Workbook").getDispatch();

		Dispatch workSheets = Dispatch.get(workBook, "Worksheets").getDispatch();

		Dispatch workSheetItem = Dispatch.call(workSheets, "Item", new Variant(1)).toDispatch();

		Dispatch cell = Dispatch.invoke(workSheetItem, "Range", Dispatch.Get, new Object[]
		{ "B3" }, new int[1]).toDispatch();
		Dispatch.put(cell, "Value", 12);
		// 7606
		// �޸ĵ���ϵ�е�ͼ������
		// Dispatch Seri1 = Dispatch.call(chartObj, "SeriesCollection", new
		// Variant(3)).toDispatch();
		// Dispatch.put(Seri1,"ChartType",4);
		updateSeriChartType(chartObj, 3, 4);
		Dispatch chartArea = Dispatch.get(chartObj, "ChartArea").toDispatch();

		// ���
		/*
		 * Dispatch tableDisp = addTable(shapes, 3, 4, 50, 130, 600, 300); Dispatch
		 * tableObj=Dispatch.get(tableDisp, "Table").getDispatch();
		 * 
		 * //�޸ı����ʽ Dispatch.invoke(tableObj, "ApplyStyle", 1, new
		 * Object[]{"{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"}, new int[1]); //��ȡ��Ԫ��
		 * Dispatch cell = getCellOfTable(tableObj, 1, 1); Dispatch cell2 =
		 * getCellOfTable(tableObj, 2, 1); addCellValue(cell, "���������");
		 * addCellValue(cell2, "����"); //�ϲ���Ԫ�� //mergeCell(cell, cell2);
		 * addTableColumn(tableObj, 4); addTableRow(tableObj, 3);
		 */
		// powerpoint�õ�չʾ���ö���
		ActiveXComponent setting = presentation.getPropertyAsComponent("SlideShowSettings");
		// setting.invoke("Run");

		// ����ppt
		presentation.invoke("SaveAs", new Variant("d:/a.ppt"));
		// �ͷſ����߳�
		ComThread.Release();
		// closePpt();
	}

	/**
	 * �������е�PPT
	 * 
	 * @param filePath
	 * @throws Exception
	 */
	public void invokePPTTemplate(String filePath) throws Exception
	{
		Dispatch windows = presentation.getProperty("Windows").toDispatch();
		Dispatch window = Dispatch.call(windows, "Item", new Variant(1)).toDispatch();
		// ��õڼ���PPT
		Dispatch pptPage = getPptPage(1);
		Dispatch selection = Dispatch.get(window, "Selection").toDispatch();
		Dispatch slideRange = Dispatch.get(selection, "SlideRange").getDispatch();
		Dispatch shapes = Dispatch.get(slideRange, "Shapes").getDispatch();
		// ��ȡ�õ�Ƭ�еĵ�N��Ԫ��
		Dispatch tableDisp = Dispatch.call(shapes, "Item", new Variant(2)).toDispatch();
		// ת��ΪTable����
		Dispatch tableObj = Dispatch.get(tableDisp, "Table").toDispatch();
		// ��ö����Style������
		Dispatch tableStyle = Dispatch.get(tableObj, "Style").toDispatch();
		System.out.println(Dispatch.get(tableStyle, "Id"));
		// {91EBBBCC-DAD2-459C-BE2E-F6DE35CF9A28}
		/*
		 * //ѡ�е�һ��Ԫ�� Dispatch.call(shape1, "Select"); Dispatch
		 * chartObj=Dispatch.get(shape1, "Chart").getDispatch();
		 * 
		 * Dispatch chartData=Dispatch.get(chartObj, "ChartData").toDispatch();
		 * 
		 * Dispatch workBook=Dispatch.get(chartData, "Workbook").getDispatch();
		 * 
		 * Dispatch workSheets=Dispatch.get(workBook, "Worksheets").getDispatch();
		 * 
		 * Dispatch workSheetItem = Dispatch.call(workSheets, "Item", new
		 * Variant(1)).toDispatch();
		 * 
		 * Dispatch cell = Dispatch.invoke(shape1, "Range", Dispatch.Get, new Object[] {
		 * "B3" }, new int[1]).toDispatch();
		 */
		// �޸ĵ���ϵ�е�ͼ������
		// Dispatch Seri1 = Dispatch.call(chartObj, "SeriesCollection", new
		// Variant(1)).toDispatch();
		// Dispatch.put(Seri1,"ChartType","4");

		// Dispatch chartArea = Dispatch.get(chartObj, "ChartArea").toDispatch();
		// ���ͼ������
		// Dispatch.call(chartArea, "ClearContents");

		// powerpoint�õ�չʾ���ö���
		ActiveXComponent setting = presentation.getPropertyAsComponent("SlideShowSettings");
		// setting.invoke("Run");
		// ����ppt
		// presentation.invoke("SaveAs", new Variant(filePath));
		// �ͷſ����߳�
		ComThread.Release();
	}
	
	/**
	 * pp��ת�����ͣ�https://msdn.microsoft.com/zh-tw/VBA/PowerPoint-VBA/articles/ppsaveasfiletype-enumeration-powerpoint
	 */
	@Override
	public void conversion(String readPath, String curFormat, String writePath, String tarFormat)
	{
		// TODO Auto-generated method stub
		String[] formatName = {"ppt", "pptx", "pdf", "bmp", "jpg", "rtf"};
		int[] formatId = {7, 11, 32, 23, 17, 6};
		Map<String, Integer> formatTo = new HashMap<String, Integer>(); 
		
		for(int i=0; i<formatName.length; i++)
		{
			formatTo.put(formatName[i], formatId[i]);
		}
		saveAs(presentation, writePath, formatTo.get(tarFormat));
	}

	@Override
	public void close()
	{
		// TODO Auto-generated method stub
		// Dispatch.call(xls, "Save");
//		if (presentation != null)
//		{
//			Dispatch.call(presentation, "Close");
//		}
//		ppt.invoke("Quit", new Variant[]{});
		ComThread.Release();
	}
}
