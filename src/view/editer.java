package view;

import javax.swing.*;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.text.BadLocationException;

import ctrl.conver;
import ctrl.fileFormat;
import ctrl.wordBean;

import java.awt.*;
import java.io.*;
import java.awt.event.*;

/**
 * 优秀的编辑器
 * 
 * @author Wenzhou
 *
 */
public class editer extends JFrame implements DocumentListener
{
	private static final long serialVersionUID = 1L;
	private JMenuBar menubar = new JMenuBar();
	private JMenu file = new JMenu("文件");
	private JMenu edit = new JMenu("编辑");
	private JMenu form = new JMenu("格式");
	private JMenu look = new JMenu("查看");
	private JMenu help = new JMenu("帮助");
	private JTextArea wordArea = new JTextArea();
	private JScrollPane imgScrollPane = new JScrollPane(wordArea);
	String[] file_items ={ "保存", "另存为", "页面设置", "打印", "退出" };
	String[] edit_items ={ "剪切", "复制", "粘贴", "查找", "替换" };
	String[] form_items ={ "自动换行", "字体" };
	private Font f1 = new Font("黑体", Font.PLAIN, 15);

	private int flag = 0;
	private String source = "";
	
	private wordBean wb = null;
	
	private String curr_path = null;
	private String targ_path = null;

//	public static void main(String[] args)
//	{
//		editer wordwin = new editer("D:\\Documents\\test.doc", "");
////		wordwin.setContent();
//		wordwin.setVisible(true);
//	}

	public editer(String cpath, String tpath)
	{
		this.curr_path = cpath;
		this.targ_path = tpath;
		try
		{
			wb = new wordBean();
			wb.openDocument(this.curr_path);
		}
		catch(Exception e)
		{
			e.printStackTrace();
			JOptionPane.showMessageDialog(null, "源文件不存在", "警告", JOptionPane.ERROR_MESSAGE);
			wb.close();
		}
		
		//初始化编辑界面
		initialize();
	}

	private void initialize()
	{
		setTitle("文本编辑器");
		Toolkit kit = Toolkit.getDefaultToolkit();
		Dimension screenSize = kit.getScreenSize();// 获取屏幕分辨率
		setSize(screenSize.width / 2, screenSize.height / 2);// 大小
		setLocation(screenSize.width / 4, screenSize.height / 4);// 位置
		add(imgScrollPane, BorderLayout.CENTER);
		setJMenuBar(menubar);
		file.setFont(f1);
		edit.setFont(f1);
		form.setFont(f1);
		look.setFont(f1);
		help.setFont(f1);
		menubar.add(file);
		menubar.add(edit);
		menubar.add(form);
		menubar.add(look);
		menubar.add(help);
		
		this.setContent();	//文本框中内容初始化
		// 添加监听器
		wordArea.getDocument().addDocumentListener(this);
		for (int i = 0; i < file_items.length; i++)
		{
			JMenuItem item1 = new JMenuItem(file_items[i]);
			item1.addActionListener(new MyActionListener1());
			item1.setFont(f1);
			file.add(item1);
		}
		for (int i = 0; i < edit_items.length; i++)
		{
			JMenuItem item2 = new JMenuItem(edit_items[i]);
			item2.addActionListener(new MyActionListener1());
			item2.setFont(f1);
			edit.add(item2);
		}
		for (int i = 0; i < form_items.length; i++)
		{
			JMenuItem item3 = new JMenuItem(form_items[i]);
			item3.addActionListener(new MyActionListener1());
			item3.setFont(f1);
			form.add(item3);
		}
//		this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		this.addWindowListener(new WindowAdapter()
		{
			public void windowClosing(WindowEvent e)
			{
				editer.this.exit();
			}
		});

	}

	public void changedUpdate(DocumentEvent e)
	{
		flag = 1;
	}

	public void insertUpdate(DocumentEvent e)
	{
		flag = 1;
		try
		{
			for(int i=0; i<e.getLength(); i++)
			{
				this.wb.moveStart();
				this.wb.moveRight(e.getOffset()+i);
				this.wb.insertText(this.wordArea.getText(e.getOffset()+i, 1));
//				System.out.println(this.wordArea.getText(e.getOffset()+i, 1));
			}
		} catch (BadLocationException e1)
		{
			e1.printStackTrace();
//			this.wb.close();
		}
//		System.out.println(wordArea.getText());
//        String inputStr=wordArea.getText().trim();
//        System.out.println("inputStr: "+ inputStr);
	}
	
	public void removeUpdate(DocumentEvent e)
	{
		flag = 1;
		try
		{
			for(int i=0; i<e.getLength(); i++)
			{
				this.wb.moveStart();
				this.wb.moveRight(e.getOffset()+1);
				this.wb.deleteText();
			}
		}
		catch(Exception e2)
		{
			e2.printStackTrace();
			this.wb.close();
		}
//		System.out.println(wb.getParagraphs(4));
	}

	/**
	 * 设置框框中的内容
	 */
	public void setContent()
	{
		int paraN = 0;
		paraN = this.wb.getParagraphsCnt();
		for(int i=1; i<=paraN; i++)
		{
			this.wordArea.append(this.wb.getParagraphs(i));
//			this.wordArea.append("\n");
//			System.out.println(this.wb.getParagraphs(i));
		}
	}

	/**
	 * 打开文件
	 */
	/*
	void open()
	{
		FileDialog filedialog = new FileDialog(this, "打开", FileDialog.LOAD);
		filedialog.setVisible(true);
		String path = filedialog.getDirectory();
		String name = filedialog.getFile();
		if (path != null && name != null)
		{
			FileInputStream file = null;
			try
			{
				file = new FileInputStream(path + name);
			} catch (FileNotFoundException e)
			{
				e.printStackTrace();
			}
			InputStreamReader put = new InputStreamReader(file);
			BufferedReader in = new BufferedReader(put);
			String temp = "";
			String now = null;
			try
			{
				now = (String) in.readLine();
			} catch (IOException e)
			{

				e.printStackTrace();
			}
			while (now != null)
			{
				temp += now + "\r\n";
				try
				{
					now = (String) in.readLine();
				} catch (IOException e)
				{

					e.printStackTrace();
				}
			}
			wordArea.setText(temp);
		}
	}
	*/

	/**
	 * 保存文件
	 */
	void save()
	{
		if(this.targ_path==null)
		{
			this.wb.save();
		}
		else
		{
			try
			{
				fileFormat ff = new fileFormat(targ_path);
				if(ff.getFormat()==0)
					this.wb.conversion(targ_path, ff.name[1]);
				else
				{
					throw new Exception();
				}
			} catch (Exception e1)
			{
				// TODO Auto-generated catch block
				e1.printStackTrace();
				JOptionPane.showMessageDialog(null, "文件转换错误", "警告", JOptionPane.ERROR_MESSAGE);
			}
		}
	}

	/**
	 * 新建文件
	 */
/*
	void newfile()
	{
		if (flag == 0)
		{
			wordArea.setText("");
		}
		if (flag == 1)
		{
			int m = JOptionPane.showConfirmDialog(this, "是否保存该文件");
			if (m == 0)
			{
				save();
				wordArea.setText("");
			}

			if (m == 1)
			{
				// System.exit(0);
				wordArea.setText("");
				flag = 0;
			}
		}
	}
*/
	/**
	 * 退出， 若未保存，询问是否保存
	 */
	void exit()
	{
		if (flag == 0)
		{
			wb.closeDocument();
			wb.close();
		}
		if (flag == 1)
		{
			int m = JOptionPane.showConfirmDialog(this, "是否保存该文件");
			if (m == 0)
			{
				save();
				wb.closeDocument();
				wb.close();
				flag = 0;
			}
			if (m == 1)
			{
				wb.closeDocument();
				wb.close();
				flag = 0;
			}
			
		}
	}

	class MyActionListener1 implements ActionListener
	{
		public void actionPerformed(ActionEvent e)
		{
			if (e.getSource() instanceof JMenuItem)
			{
				if (e.getActionCommand() == "剪切")
				{
					wordArea.cut();
				}
				if (e.getActionCommand() == "复制")
				{
					wordArea.copy();
				}
				if (e.getActionCommand() == "粘贴")
				{
					wordArea.paste();
				}
//				if (e.getActionCommand() == "打开")
//				{
//					open();
//				}
				if (e.getActionCommand() == "保存")
				{
					save();
				}
//				if (e.getActionCommand() == "新建")
//				{
//					newfile();
//				}
				if (e.getActionCommand() == "退出")
				{
					exit();
				}
				if (e.getActionCommand() == "打印")
				{
					try
					{
						editer.this.wb.printFile();
					}
					catch(Exception e1)
					{
						JOptionPane.showMessageDialog(null, "打印失败", "警告", JOptionPane.ERROR_MESSAGE);
					}
				}
			}

		}
	}

}