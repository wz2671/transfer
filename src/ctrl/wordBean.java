package ctrl;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

/**
 * jacob����MSword��
 * 
 * @author
 */

public class wordBean extends bean
{
	// word�ĵ�
	private Dispatch doc;

	// word���г������
	private ActiveXComponent word;

	// ����word�ĵ�����
	private Dispatch documents;

	// ѡ���ķ�Χ������
	private Dispatch selection;

	private boolean saveOnExit = false;

	public wordBean() throws Exception
	{
		ComThread.InitSTA();

		if (word == null)
		{
			word = new ActiveXComponent("Word.Application");
			word.setProperty("Visible", new Variant(false)); // ���ɼ���word
			word.setProperty("AutomationSecurity", new Variant(3)); // ���ú�
		}
		if (documents == null)
			documents = word.getProperty("Documents").toDispatch();
	}
 
	/**
	 * �����˳�ʱ����
	 * 
	 * @param saveOnExit
	 *  	  boolean true-�˳�ʱ�����ļ���false-�˳�ʱ�������ļ�
	 */
	public void setSaveOnExit(boolean saveOnExit)
	{
		this.saveOnExit = saveOnExit;
	}

	/**
	 * ����һ���µ�word�ĵ�
	 * 
	 */
	public void createNewDocument()
	{
		doc = Dispatch.call(documents, "Add").toDispatch();
		selection = Dispatch.get(word, "Selection").toDispatch();
	}

	/**
	 * ��һ���Ѵ��ڵ��ĵ�
	 * 
	 * @param docPath
	 */
	public void openDocument(String docPath)
	{
		closeDocument();
		doc = Dispatch.call(documents, "Open", docPath).toDispatch();
		selection = Dispatch.get(word, "Selection").toDispatch();
	}

	/**
	 * ��һ�������ĵ�,
	 * 
	 * @param docPath-�ļ�ȫ��
	 * @param pwd-����
	 */
	public void openDocumentOnlyRead(String docPath, String pwd) throws Exception
	{
		closeDocument();
		// doc = Dispatch.invoke(documents, "Open", Dispatch.Method,
		// new Object[]{docPath, new Variant(false), new Variant(true), new
		// Variant(true), pwd},
		// new int[1]).toDispatch();//��word�ļ�
		doc = Dispatch.callN(documents, "Open", new Object[]
		{ docPath, new Variant(false), new Variant(true), new Variant(true), pwd, "", new Variant(false) })
				.toDispatch();
		selection = Dispatch.get(word, "Selection").toDispatch();
	}

	public void openDocument(String docPath, String pwd) throws Exception
	{
		closeDocument();
		doc = Dispatch.callN(documents, "Open", new Object[]
		{ docPath, new Variant(false), new Variant(false), new Variant(true), pwd }).toDispatch();
		selection = Dispatch.get(word, "Selection").toDispatch();
	}

	/**
	 * ��ѡ�������ݻ����������ƶ�
	 * 
	 * @param pos �ƶ��ľ���
	 */
	public void moveUp(int pos)
	{
		if (selection == null)
			selection = Dispatch.get(word, "Selection").toDispatch();
		for (int i = 0; i < pos; i++)
			Dispatch.call(selection, "MoveUp");
	}

	/**
	 * ��ѡ�������ݻ��߲���������ƶ�
	 * 
	 * @param pos �ƶ��ľ���
	 */
	public void moveDown(int pos)
	{
		if (selection == null)
			selection = Dispatch.get(word, "Selection").toDispatch();
		for (int i = 0; i < pos; i++)
			Dispatch.call(selection, "MoveDown");
	}

	/**
	 * ��ѡ�������ݻ��߲���������ƶ�
	 * 
	 * @param pos �ƶ��ľ���
	 */
	public void moveLeft(int pos)
	{
		if (selection == null)
			selection = Dispatch.get(word, "Selection").toDispatch();
		for (int i = 0; i < pos; i++)
		{
			Dispatch.call(selection, "MoveLeft");
		}
	}

	/**
	 * ��ѡ�������ݻ��߲���������ƶ�
	 * 
	 * @param pos �ƶ��ľ���
	 */
	public void moveRight(int pos)
	{
		if (selection == null)
			selection = Dispatch.get(word, "Selection").toDispatch();
		for (int i = 0; i < pos; i++)
			Dispatch.call(selection, "MoveRight");
	}

	/**
	 * �Ѳ�����ƶ����ļ���λ��
	 * 
	 */
	public void moveStart()
	{
		if (selection == null)
			selection = Dispatch.get(word, "Selection").toDispatch();
		Dispatch.call(selection, "HomeKey", new Variant(6));
	}

	/**
	 * ��ѡ�����ݻ����㿪ʼ�����ı�
	 * 
	 * @param toFindText Ҫ���ҵ��ı�
	 * @return boolean true-���ҵ���ѡ�и��ı���false-δ���ҵ��ı�
	 */
	@SuppressWarnings("static-access")
	public boolean find(String toFindText)
	{
		if (toFindText == null || toFindText.equals(""))
			return false;
		// ��selection����λ�ÿ�ʼ��ѯ
		Dispatch find = word.call(selection, "Find").toDispatch();
		// ����Ҫ���ҵ�����
		Dispatch.put(find, "Text", toFindText);
		// ��ǰ����
		Dispatch.put(find, "Forward", "True");
		// ���ø�ʽ
		Dispatch.put(find, "Format", "True");
		// ��Сдƥ��
		Dispatch.put(find, "MatchCase", "True");
		// ȫ��ƥ��
		Dispatch.put(find, "MatchWholeWord", "True");
		// ���Ҳ�ѡ��
		return Dispatch.call(find, "Execute").getBoolean();
	}

	/**
	 * ��ѡ��ѡ�������趨Ϊ�滻�ı�
	 * 
	 * @param toFindText �����ַ���
	 * @param newText Ҫ�滻������
	 * @return
	 */
	public boolean replaceText(String toFindText, String newText)
	{
		if (!find(toFindText))
			return false;
		Dispatch.put(selection, "Text", newText);
		return true;
	}

	/**
	 * ȫ���滻�ı�
	 * 
	 * @param toFindText �����ַ���
	 * @param newText Ҫ�滻������
	 */
	public void replaceAllText(String toFindText, String newText)
	{
		while (find(toFindText))
		{
			Dispatch.put(selection, "Text", newText);
			Dispatch.call(selection, "MoveRight");
		}
	}

	/**
	 * �ڵ�ǰ���������ַ���
	 * 
	 * @param newText Ҫ��������ַ���
	 */
	public void insertText(String newText)
	{
		Dispatch.put(selection, "Text", newText);
	}
	
	/**
	 * �˸�
	 */
	public void deleteText()
	{
		Dispatch.call(selection, "TypeBackspace");
	}

	/**
	 * 
	 * @param toFindText Ҫ���ҵ��ַ���
	 * @param imagePath ͼƬ·��
	 * @return
	 */
	public boolean replaceImage(String toFindText, String imagePath)
	{
		if (!find(toFindText))
			return false;
		Dispatch.call(Dispatch.get(selection, "InLineShapes").toDispatch(), "AddPicture", imagePath);
		return true;
	}

	/**
	 * ȫ���滻ͼƬ
	 * 
	 * @param toFindText �����ַ���
	 * @param imagePath ͼƬ·��
	 */
	public void replaceAllImage(String toFindText, String imagePath)
	{
		while (find(toFindText))
		{
			Dispatch.call(Dispatch.get(selection, "InLineShapes").toDispatch(), "AddPicture", imagePath);
			Dispatch.call(selection, "MoveRight");
		}
	}

	/**
	 * �ڵ�ǰ��������ͼƬ
	 * 
	 * @param imagePath ͼƬ·��
	 */
	public void insertImage(String imagePath)
	{
		Dispatch.call(Dispatch.get(selection, "InLineShapes").toDispatch(), "AddPicture", imagePath);
	}

	/**
	 * �ϲ���Ԫ��
	 * 
	 * @param tableIndex
	 * @param fstCellRowIdx
	 * @param fstCellColIdx
	 * @param secCellRowIdx
	 * @param secCellColIdx
	 */
	public void mergeCell(int tableIndex, int fstCellRowIdx, int fstCellColIdx, int secCellRowIdx, int secCellColIdx)
	{
		// ���б��
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// Ҫ���ı��
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		Dispatch fstCell = Dispatch.call(table, "Cell", new Variant(fstCellRowIdx), new Variant(fstCellColIdx))
				.toDispatch();
		Dispatch secCell = Dispatch.call(table, "Cell", new Variant(secCellRowIdx), new Variant(secCellColIdx))
				.toDispatch();
		Dispatch.call(fstCell, "Merge", secCell);
	}

	/**
	 * ��ָ���ĵ�Ԫ������д����
	 * 
	 * @param tableIndex
	 * @param cellRowIdx
	 * @param cellColIdx
	 * @param txt
	 */
	public void putTxtToCell(int tableIndex, int cellRowIdx, int cellColIdx, String txt)
	{
		// ���б��
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// Ҫ���ı��
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		Dispatch cell = Dispatch.call(table, "Cell", new Variant(cellRowIdx), new Variant(cellColIdx)).toDispatch();
		Dispatch.call(cell, "Select");
		Dispatch.put(selection, "Text", txt);
	}

	/**
	 * ���ָ���ĵ�Ԫ��������
	 * 
	 * @param tableIndex
	 * @param cellRowIdx
	 * @param cellColIdx
	 * @return
	 */
	public String getTxtFromCell(int tableIndex, int cellRowIdx, int cellColIdx)
	{
		// ���б��
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// Ҫ���ı��
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		Dispatch cell = Dispatch.call(table, "Cell", new Variant(cellRowIdx), new Variant(cellColIdx)).toDispatch();
		Dispatch.call(cell, "Select");
		String ret = "";
		ret = Dispatch.get(selection, "Text").toString();
		ret = ret.substring(0, ret.length() - 1); // ȥ�����Ļس���;
		return ret;
	}

	/**
	 * �ڵ�ǰ�ĵ���������������
	 * 
	 * @param po
	 */
	public void pasteExcelSheet(String pos)
	{
		moveStart();
		if (this.find(pos))
		{
			Dispatch textRange = Dispatch.get(selection, "Range").toDispatch();
			Dispatch.call(textRange, "Paste");
		}
	}

	/**
	 * �ڵ�ǰ�ĵ�ָ����λ�ÿ������
	 * 
	 * @param pos ��ǰ�ĵ�ָ����λ��
	 * @param tableIndex �������ı����word�ĵ���������λ��
	 */
	public void copyTable(String pos, int tableIndex)
	{
		// ���б��
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// Ҫ���ı��
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		Dispatch range = Dispatch.get(table, "Range").toDispatch();
		Dispatch.call(range, "Copy");
		if (this.find(pos))
		{
			Dispatch textRange = Dispatch.get(selection, "Range").toDispatch();
			Dispatch.call(textRange, "Paste");
		}
	}

	/**
	 * �ڵ�ǰ�ĵ�ָ����λ�ÿ���������һ���ĵ��еı��
	 * 
	 * @param anotherDocPath ��һ���ĵ��Ĵ���·��
	 * @param tableIndex �������ı������һ���ĵ��е�λ��
	 * @param pos ��ǰ�ĵ�ָ����λ��
	 */
	public void copyTableFromAnotherDoc(String anotherDocPath, int tableIndex, String pos)
	{
		Dispatch doc2 = null;
		try
		{
			doc2 = Dispatch.call(documents, "Open", anotherDocPath).toDispatch();
			// ���б��
			Dispatch tables = Dispatch.get(doc2, "Tables").toDispatch();
			// Ҫ���ı��
			Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
			Dispatch range = Dispatch.get(table, "Range").toDispatch();
			Dispatch.call(range, "Copy");
			if (this.find(pos))
			{
				Dispatch textRange = Dispatch.get(selection, "Range").toDispatch();
				Dispatch.call(textRange, "Paste");
			}
		} catch (Exception e)
		{
			e.printStackTrace();
		} finally
		{
			if (doc2 != null)
			{
				Dispatch.call(doc2, "Close", new Variant(saveOnExit));
				doc2 = null;
			}
		}
	}

	/**
	 * �ڵ�ǰ�ĵ�ָ����λ�ÿ���������һ���ĵ��е�ͼƬ
	 * 
	 * @param anotherDocPath ��һ���ĵ��Ĵ���·��
	 * @param shapeIndex ��������ͼƬ����һ���ĵ��е�λ��
	 * @param pos ��ǰ�ĵ�ָ����λ��
	 */
	public void copyImageFromAnotherDoc(String anotherDocPath, int shapeIndex, String pos)
	{
		Dispatch doc2 = null;
		try
		{
			doc2 = Dispatch.call(documents, "Open", anotherDocPath).toDispatch();
			Dispatch shapes = Dispatch.get(doc2, "InLineShapes").toDispatch();
			Dispatch shape = Dispatch.call(shapes, "Item", new Variant(shapeIndex)).toDispatch();
			Dispatch imageRange = Dispatch.get(shape, "Range").toDispatch();
			Dispatch.call(imageRange, "Copy");
			if (this.find(pos))
			{
				Dispatch textRange = Dispatch.get(selection, "Range").toDispatch();
				Dispatch.call(textRange, "Paste");
			}
		} catch (Exception e)
		{
			e.printStackTrace();
		} finally
		{
			if (doc2 != null)
			{
				Dispatch.call(doc2, "Close", new Variant(saveOnExit));
				doc2 = null;
			}
		}
	}

	/**
	 * �������
	 * 
	 * @param pos λ��
	 * @param cols ����
	 * @param rows ����
	 */
	public void createTable(String pos, int numCols, int numRows)
	{
		if (find(pos))
		{
			Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
			Dispatch range = Dispatch.get(selection, "Range").toDispatch();
			@SuppressWarnings("unused")
			Dispatch newTable = Dispatch.call(tables, "Add", range, new Variant(numRows), new Variant(numCols))
					.toDispatch();
			Dispatch.call(selection, "MoveRight");
		} else
		{
			Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
			Dispatch range = Dispatch.get(selection, "Range").toDispatch();
			@SuppressWarnings("unused")
			Dispatch newTable = Dispatch.call(tables, "Add", range, new Variant(numRows), new Variant(numCols))
					.toDispatch();
			Dispatch.call(selection, "MoveRight");
		}
	}

	/**
	 * ��ָ����ǰ��������
	 * 
	 * @param tableIndex word�ļ��еĵ�N�ű�(��1��ʼ)
	 * @param rowIndex ָ���е����(��1��ʼ)
	 */
	public void addTableRow(int tableIndex, int rowIndex)
	{
		// ���б��
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// Ҫ���ı��
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		// ����������
		Dispatch rows = Dispatch.get(table, "Rows").toDispatch();
		Dispatch row = Dispatch.call(rows, "Item", new Variant(rowIndex)).toDispatch();
		Dispatch.call(rows, "Add", new Variant(row));
	}

	/**
	 * �ڵ�1��ǰ����һ��
	 * 
	 * @param tableIndex word�ĵ��еĵ�N�ű�(��1��ʼ)
	 */
	public void addFirstTableRow(int tableIndex)
	{
		// ���б��
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// Ҫ���ı��
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		// ����������
		Dispatch rows = Dispatch.get(table, "Rows").toDispatch();
		Dispatch row = Dispatch.get(rows, "First").toDispatch();
		Dispatch.call(rows, "Add", new Variant(row));
	}

	/**
	 * �����1��ǰ����һ��
	 * 
	 * @param tableIndex word�ĵ��еĵ�N�ű�(��1��ʼ)
	 */
	public void addLastTableRow(int tableIndex)
	{
		// ���б��
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// Ҫ���ı��
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		// ����������
		Dispatch rows = Dispatch.get(table, "Rows").toDispatch();
		Dispatch row = Dispatch.get(rows, "Last").toDispatch();
		Dispatch.call(rows, "Add", new Variant(row));
	}

	/**
	 * ����һ��
	 * 
	 * @param tableIndex word�ĵ��еĵ�N�ű�(��1��ʼ)
	 */
	public void addRow(int tableIndex)
	{
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// Ҫ���ı��
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		// ����������
		Dispatch rows = Dispatch.get(table, "Rows").toDispatch();
		Dispatch.call(rows, "Add");
	}

	/**
	 * ����һ��
	 * 
	 * @param tableIndex  word�ĵ��еĵ�N�ű�(��1��ʼ)
	 */
	public void addCol(int tableIndex)
	{
		// ���б��
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// Ҫ���ı��
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		// ����������
		Dispatch cols = Dispatch.get(table, "Columns").toDispatch();
		Dispatch.call(cols, "Add").toDispatch();
		Dispatch.call(cols, "AutoFit");
	}

	/**
	 * ��ָ����ǰ�����ӱ�����
	 * 
	 * @param tableIndex word�ĵ��еĵ�N�ű�(��1��ʼ)
	 * @param colIndex �ƶ��е���� (��1��ʼ)
	 */
	public void addTableCol(int tableIndex, int colIndex)
	{
		// ���б��
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// Ҫ���ı��
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		// ����������
		Dispatch cols = Dispatch.get(table, "Columns").toDispatch();
		System.out.println(Dispatch.get(cols, "Count"));
		Dispatch col = Dispatch.call(cols, "Item", new Variant(colIndex)).toDispatch();
		// Dispatch col = Dispatch.get(cols, "First").toDispatch();
		Dispatch.call(cols, "Add", col).toDispatch();
		Dispatch.call(cols, "AutoFit");
	}

	/**
	 * �ڵ�1��ǰ����һ��
	 * 
	 * @param tableIndex word�ĵ��еĵ�N�ű�(��1��ʼ)
	 */
	public void addFirstTableCol(int tableIndex)
	{
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// Ҫ���ı��
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		// ����������
		Dispatch cols = Dispatch.get(table, "Columns").toDispatch();
		Dispatch col = Dispatch.get(cols, "First").toDispatch();
		Dispatch.call(cols, "Add", col).toDispatch();
		Dispatch.call(cols, "AutoFit");
	}

	/**
	 * �����һ��ǰ����һ��
	 * 
	 * @param tableIndex word�ĵ��еĵ�N�ű�(��1��ʼ)
	 */
	public void addLastTableCol(int tableIndex)
	{
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// Ҫ���ı��
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		// ����������
		Dispatch cols = Dispatch.get(table, "Columns").toDispatch();
		Dispatch col = Dispatch.get(cols, "Last").toDispatch();
		Dispatch.call(cols, "Add", col).toDispatch();
		Dispatch.call(cols, "AutoFit");
	}

	/**
	 * �Զ��������
	 *
	 */
	@SuppressWarnings("deprecation")
	public void autoFitTable()
	{
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		int count = Dispatch.get(tables, "Count").toInt();
		for (int i = 0; i < count; i++)
		{
			Dispatch table = Dispatch.call(tables, "Item", new Variant(i + 1)).toDispatch();
			Dispatch cols = Dispatch.get(table, "Columns").toDispatch();
			Dispatch.call(cols, "AutoFit");
		}
	}

	/**
	 * ����word��ĺ��Ե������Ŀ��,���к걣����document��
	 *
	 */
	@SuppressWarnings("deprecation")
	public void callWordMacro()
	{
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		int count = Dispatch.get(tables, "Count").toInt();
		Variant vMacroName = new Variant("Normal.NewMacros.tableFit");
		@SuppressWarnings("unused")
		Variant vParam = new Variant("param1");
		@SuppressWarnings("unused")
		Variant para[] = new Variant[]
		{ vMacroName };
		for (int i = 0; i < count; i++)
		{
			Dispatch table = Dispatch.call(tables, "Item", new Variant(i + 1)).toDispatch();
			Dispatch.call(table, "Select");
			Dispatch.call(word, "Run", "tableFitContent");
		}
	}

	/**
	 * ���õ�ǰѡ�����ݵ�����
	 * 
	 * @param boldSize
	 * @param italicSize
	 * @param underLineSize �»���
	 * @param colorSize ������ɫ
	 * @param size �����С
	 * @param name ��������
	 */
	public void setFont(boolean bold, boolean italic, boolean underLine, String colorSize, String size, String name)
	{
		Dispatch font = Dispatch.get(selection, "Font").toDispatch();
		Dispatch.put(font, "Name", new Variant(name));
		Dispatch.put(font, "Bold", new Variant(bold));
		Dispatch.put(font, "Italic", new Variant(italic));
		Dispatch.put(font, "Underline", new Variant(underLine));
		Dispatch.put(font, "Color", colorSize);
		Dispatch.put(font, "Size", size);
	}

	/**
	 * ���õ�Ԫ��ѡ��
	 * 
	 * @param tableIndex
	 * @param cellRowIdx
	 * @param cellColIdx
	 */
	public void setTableCellSelected(int tableIndex, int cellRowIdx, int cellColIdx)
	{
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();
		Dispatch cell = Dispatch.call(table, "Cell", new Variant(cellRowIdx), new Variant(cellColIdx)).toDispatch();
		Dispatch.call(cell, "Select");
	}

	/**
	 * ����ѡ����Ԫ��Ĵ�ֱ����ʽ, ��ʹ��setTableCellSelectedѡ��һ����Ԫ��
	 * 
	 * @param align 0-����, 1-����, 3-�׶�
	 */
	public void setCellVerticalAlign(int verticalAlign)
	{
		Dispatch cells = Dispatch.get(selection, "Cells").toDispatch();
		Dispatch.put(cells, "VerticalAlignment", new Variant(verticalAlign));
	}

	/**
	 * ���õ�ǰ�ĵ������б��ˮƽ���з�ʽ������һЩ��ʽ,���ڽ�word�ļ�ת��Ϊhtml��,����걨��
	 */
	@SuppressWarnings("deprecation")
	public void setApplyTableFormat()
	{
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		int tabCount = Integer.valueOf(Dispatch.get(tables, "Count").toString()); // System.out.println(tabCount);
		System.out.println("*******************************************************");
		for (int i = 1; i <= tabCount; i++)
		{
			Dispatch table = Dispatch.call(tables, "Item", new Variant(i)).toDispatch();
			Dispatch rows = Dispatch.get(table, "Rows").toDispatch();

			if (i == 1)
			{
				Dispatch.put(rows, "Alignment", new Variant(2)); // 1-����,2-Right
				continue;
			}
			Dispatch.put(rows, "Alignment", new Variant(1)); // 1-����
			Dispatch.call(table, "AutoFitBehavior", new Variant(1));// �����Զ��������ʽ,1-���ݴ����Զ�����
			Dispatch.put(table, "PreferredWidthType", new Variant(1));
			Dispatch.put(table, "PreferredWidth", new Variant(700));
			System.out.println(Dispatch.get(rows, "HeightRule").toString());
			Dispatch.put(rows, "HeightRule", new Variant(1)); // 0-�Զ�wdRowHeightAuto,1-��СֵwdRowHeightAtLeast,
																// 2-�̶�wdRowHeightExactly
			Dispatch.put(rows, "Height", new Variant(0.04 * 28.35));
			// int oldAlign = Integer.valueOf(Dispatch.get(rows, "Alignment").toString());
			// System.out.println("Algin:" + oldAlign);
		}
	}

	/**
	 * ���ö����ʽ
	 * 
	 * @param alignment 0-�����, 1-�Ҷ���, 2-�Ҷ���, 3-���˶���, 4-��ɢ����
	 * @param lineSpaceingRule
	 * @param lineUnitBefore
	 * @param lineUnitAfter
	 * @param characterUnitFirstLineIndent
	 */
	public void setParagraphsProperties(int alignment, int lineSpaceingRule, int lineUnitBefore, int lineUnitAfter,
			int characterUnitFirstLineIndent)
	{
		Dispatch paragraphs = Dispatch.get(selection, "Paragraphs").toDispatch();
		Dispatch.put(paragraphs, "Alignment", new Variant(alignment)); // ���뷽ʽ
		Dispatch.put(paragraphs, "LineSpacingRule", new Variant(lineSpaceingRule)); // �о�
		Dispatch.put(paragraphs, "LineUnitBefore", new Variant(lineUnitBefore)); // ��ǰ
		Dispatch.put(paragraphs, "LineUnitAfter", new Variant(lineUnitAfter)); // �κ�
		Dispatch.put(paragraphs, "CharacterUnitFirstLineIndent", new Variant(characterUnitFirstLineIndent)); // ���������ַ���
	}

	/**
	 * ���õ�ǰ�����ʽ, ʹ��ǰ,����ѡ�ж���
	 */
	public void getParagraphsProperties()
	{
		Dispatch paragraphs = Dispatch.get(selection, "Paragraphs").toDispatch();
		String val = Dispatch.get(paragraphs, "LineSpacingRule").toString(); // �о�
		val = Dispatch.get(paragraphs, "Alignment").toString(); // ���뷽ʽ
		val = Dispatch.get(paragraphs, "LineUnitBefore").toString(); // ��ǰ����
		val = Dispatch.get(paragraphs, "LineUnitAfter").toString(); // �κ�����
		val = Dispatch.get(paragraphs, "FirstLineIndent").toString(); // ��������
		val = Dispatch.get(paragraphs, "CharacterUnitFirstLineIndent").toString(); // ���������ַ���
	}

	/**
	 * �ļ���������Ϊ
	 * 
	 * @param savePath ��������Ϊ·��
	 */
	public void save(String savePath)
	{
		Dispatch.call(Dispatch.call(word, "WordBasic").getDispatch(), "FileSaveAs", savePath);
	}
	
	/**
	 * 
	 */
	public void save()
	{
		Dispatch.call(doc, "Save");
	}

	/**
	 * �ļ�����Ϊhtml��ʽ
	 * 
	 * @param savePath
	 * @param htmlPath
	 */
	public void saveAsHtml(String htmlPath)
	{
		Dispatch.invoke(doc, "SaveAs", Dispatch.Method, new Object[]
		{ htmlPath, new Variant(8) }, new int[1]);
	}

	/**
	 * �ر��ĵ�
	 * 
	 * @param val 0�������޸� -1 �����޸� -2 ��ʾ�Ƿ񱣴��޸�
	 */
	public void closeDocument(int val)
	{
		Dispatch.call(doc, "Close", new Variant(val));
		doc = null;
	}

	/**
	 * �رյ�ǰword�ĵ����ȱ���
	 * 
	 */
	public void closeDocument()
	{
		if (doc != null)
		{
//			Dispatch.call(doc, "Save");
			Dispatch.call(doc, "Close", new Variant(saveOnExit));
			doc = null;
		}
	}

	/*
	 * �����棬�رյ�ǰ�ĵ�
	 */
	public void closeDocumentWithoutSave()
	{
		if (doc != null)
		{
			Dispatch.call(doc, "Close", new Variant(false));
			doc = null;
		}
	}

	/**
	 * �ر�ȫ��Ӧ��
	 * 
	 */
	public void close()
	{
		// closeDocument();
		if (word != null)
		{
			Dispatch.call(word, "Quit");
			word = null;
		}
		selection = null;
		documents = null;
		ComThread.Release();
	}

	/**
	 * ��ӡ��ǰword�ĵ�
	 * 
	 */
	public void printFile()
	{
		if (doc != null)
		{
			Dispatch.call(doc, "PrintOut");
		}
	}

	/**
	 * ������ǰ��,���������, ʹ��expression.Protect(Type, NoReset, Password)
	 * 
	 * @param pwd WdProtectionType ���������� WdProtectionType ����֮һ��
	 *            1-wdAllowOnlyComments, 2-wdAllowOnlyFormFields,
	 *            0-wdAllowOnlyRevisions, -1-wdNoProtection, 3-wdAllowOnlyReading
	 * 
	 *            ʹ�ò��� main1()
	 */
	public void protectedWord(String pwd)
	{
		String protectionType = Dispatch.get(doc, "ProtectionType").toString();
		if (protectionType.equals("-1"))
		{
			Dispatch.call(doc, "Protect", new Variant(3), new Variant(true), pwd);
		}
	}

	/**
	 * ����ĵ�����,�������
	 * 
	 * @param pwd WdProtectionType ����֮һ(Long ���ͣ�ֻ��)��
	 *            1-wdAllowOnlyComments,2-wdAllowOnlyFormFields��
	 *            0-wdAllowOnlyRevisions,-1-wdNoProtection, 3-wdAllowOnlyReading
	 * 
	 *            ʹ�ò��� main1()
	 */
	public void unProtectedWord(String pwd)
	{
		String protectionType = Dispatch.get(doc, "ProtectionType").toString();
		if (protectionType.equals("3"))
		{
			Dispatch.call(doc, "Unprotect", pwd);
		}
	}

	/**
	 * ����word�ĵ���ȫ����
	 * 
	 * @param value
	 *            1-msoAutomationSecurityByUI ʹ�á���ȫ���Ի���ָ���İ�ȫ���á�
	 *            2-msoAutomationSecurityForceDisable �ڳ���򿪵������ļ��н������к꣬������ʾ�κΰ�ȫ���ѡ�
	 *            3-msoAutomationSecurityLow �������к꣬��������Ӧ�ó���ʱ��Ĭ��ֵ��
	 */
	public void setAutomationSecurity(int value)
	{
		word.setProperty("AutomationSecurity", new Variant(value));
	}

	/**
	 * ��ȡ�ĵ��е�paragraphsIndex�����ֵ�����;
	 * 
	 * @param paragraphsIndex
	 * @return
	 */
	public String getParagraphs(int paragraphsIndex)
	{
		String ret = "";
		Dispatch paragraphs = Dispatch.get(doc, "Paragraphs").toDispatch(); // ���ж���
		int paragraphCount = Dispatch.get(paragraphs, "Count").getInt(); // һ���Ķ�����
		Dispatch paragraph = null;
		Dispatch range = null;
		if (paragraphCount >= paragraphsIndex && 0 < paragraphsIndex)
		{
			paragraph = Dispatch.call(paragraphs, "Item", new Variant(paragraphsIndex)).toDispatch();
			range = Dispatch.get(paragraph, "Range").toDispatch();
			ret = Dispatch.get(range, "Text").toString();
			ret = ret.replace('\r', '\n');
		}
		return ret;
	}

	public int getParagraphsCnt()
	{
		Dispatch paragraphs = Dispatch.get(doc, "Paragraphs").toDispatch(); // ���ж���
		int paragraphCount = Dispatch.get(paragraphs, "Count").getInt();
		return paragraphCount;
	}

	/**
	 * ����ҳü����
	 * 
	 * @param cont
	 * @return
	 * 
	 * 		Sub AddHeaderText() '����ҳü��ҳ���е����� '�� Headers��Footers �� HeaderFooter
	 *         ���Է��� HeaderFooter ��������ʾ�����ĵ�ǰҳü�е����֡� With
	 *         ActiveDocument.ActiveWindow.View .SeekView = wdSeekCurrentPageHeader
	 *         Selection.HeaderFooter.Range.Text = "Header text" .SeekView =
	 *         wdSeekMainDocument End With End Sub
	 */
	public void setHeaderContent(String cont)
	{
		Dispatch activeWindow = Dispatch.get(doc, "ActiveWindow").toDispatch();
		Dispatch view = Dispatch.get(activeWindow, "View").toDispatch();
		// Dispatch seekView = Dispatch.get(view, "SeekView").toDispatch();
		Dispatch.put(view, "SeekView", new Variant(9)); // wdSeekCurrentPageHeader-9

		Dispatch headerFooter = Dispatch.get(selection, "HeaderFooter").toDispatch();
		Dispatch range = Dispatch.get(headerFooter, "Range").toDispatch();
		Dispatch.put(range, "Text", new Variant(cont));
		// String content = Dispatch.get(range, "Text").toString();
		Dispatch font = Dispatch.get(range, "Font").toDispatch();

		Dispatch.put(font, "Name", new Variant("����_GB2312"));
		Dispatch.put(font, "Bold", new Variant(true));
		// Dispatch.put(font, "Italic", new Variant(true));
		// Dispatch.put(font, "Underline", new Variant(true));
		Dispatch.put(font, "Size", 9);

		Dispatch.put(view, "SeekView", new Variant(0)); // wdSeekMainDocument-0�ָ���ͼ;
	}

//	public static void main(String[] args) throws Exception
//	{
//		wordBean word = new wordBean();
//
////		word.setHeaderContent("*****************88����ҳü����11111111111111111!");
//		
////		System.out.print(word.getParagraphs(2));
//		word.moveStart();
//		word.moveRight(10);
//		word.deleteText();
////		word.closeDocument(0);
////		word.close();
//	}

	@Override
	public void conversion(String readPath, String curFormat, String writePath, String tarFormat)
	{
		// TODO Auto-generated method stub
		String[] formatName = {"doc", "docx", "pdf", "html", "xml", "txt", "rtf"};
		int[] formatId = {0, 12, 17, 8, 19, 2, 6};
		Map<String, Integer> formatTo = new HashMap<String, Integer>(); 
		
		for(int i=0; i<formatName.length; i++)
		{
			formatTo.put(formatName[i], formatId[i]);
		}
		this.openDocument(readPath);
		Dispatch.invoke(doc, "SaveAs", Dispatch.Method, new Object[]
				{ writePath, new Variant(formatTo.get(tarFormat)) }, new int[1]);
	}
	
	public void conversion(String writePath, String tarFormat)
	{
		// TODO Auto-generated method stub
		String[] formatName = {"doc", "docx", "pdf", "html", "xml", "txt", "rtf"};
		int[] formatId = {0, 12, 17, 8, 19, 2, 6};
		Map<String, Integer> formatTo = new HashMap<String, Integer>(); 
		
		for(int i=0; i<formatName.length; i++)
		{
			formatTo.put(formatName[i], formatId[i]);
		}
		Dispatch.invoke(doc, "SaveAs", Dispatch.Method, new Object[]
				{ writePath, new Variant(formatTo.get(tarFormat)) }, new int[1]);
	}
}