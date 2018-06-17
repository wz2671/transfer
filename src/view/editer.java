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
 * ����ı༭��
 * 
 * @author Wenzhou
 *
 */
public class editer extends JFrame implements DocumentListener
{
	private static final long serialVersionUID = 1L;
	private JMenuBar menubar = new JMenuBar();
	private JMenu file = new JMenu("�ļ�");
	private JMenu edit = new JMenu("�༭");
	private JMenu form = new JMenu("��ʽ");
	private JMenu look = new JMenu("�鿴");
	private JMenu help = new JMenu("����");
	private JTextArea wordArea = new JTextArea();
	private JScrollPane imgScrollPane = new JScrollPane(wordArea);
	String[] file_items ={ "����", "���Ϊ", "ҳ������", "��ӡ", "�˳�" };
	String[] edit_items ={ "����", "����", "ճ��", "����", "�滻" };
	String[] form_items ={ "�Զ�����", "����" };
	private Font f1 = new Font("����", Font.PLAIN, 15);

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
			JOptionPane.showMessageDialog(null, "Դ�ļ�������", "����", JOptionPane.ERROR_MESSAGE);
			wb.close();
		}
		
		//��ʼ���༭����
		initialize();
	}

	private void initialize()
	{
		setTitle("�ı��༭��");
		Toolkit kit = Toolkit.getDefaultToolkit();
		Dimension screenSize = kit.getScreenSize();// ��ȡ��Ļ�ֱ���
		setSize(screenSize.width / 2, screenSize.height / 2);// ��С
		setLocation(screenSize.width / 4, screenSize.height / 4);// λ��
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
		
		this.setContent();	//�ı��������ݳ�ʼ��
		// ��Ӽ�����
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
	 * ���ÿ���е�����
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
	 * ���ļ�
	 */
	/*
	void open()
	{
		FileDialog filedialog = new FileDialog(this, "��", FileDialog.LOAD);
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
	 * �����ļ�
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
				JOptionPane.showMessageDialog(null, "�ļ�ת������", "����", JOptionPane.ERROR_MESSAGE);
			}
		}
	}

	/**
	 * �½��ļ�
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
			int m = JOptionPane.showConfirmDialog(this, "�Ƿ񱣴���ļ�");
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
	 * �˳��� ��δ���棬ѯ���Ƿ񱣴�
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
			int m = JOptionPane.showConfirmDialog(this, "�Ƿ񱣴���ļ�");
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
				if (e.getActionCommand() == "����")
				{
					wordArea.cut();
				}
				if (e.getActionCommand() == "����")
				{
					wordArea.copy();
				}
				if (e.getActionCommand() == "ճ��")
				{
					wordArea.paste();
				}
//				if (e.getActionCommand() == "��")
//				{
//					open();
//				}
				if (e.getActionCommand() == "����")
				{
					save();
				}
//				if (e.getActionCommand() == "�½�")
//				{
//					newfile();
//				}
				if (e.getActionCommand() == "�˳�")
				{
					exit();
				}
				if (e.getActionCommand() == "��ӡ")
				{
					try
					{
						editer.this.wb.printFile();
					}
					catch(Exception e1)
					{
						JOptionPane.showMessageDialog(null, "��ӡʧ��", "����", JOptionPane.ERROR_MESSAGE);
					}
				}
			}

		}
	}

}