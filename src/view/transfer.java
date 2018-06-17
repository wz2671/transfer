package view;

import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;

import java.awt.Font;

import javax.swing.JTextField;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.JButton;

import java.awt.Color;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.awt.event.ActionEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;

import javax.swing.JFileChooser;

import ctrl.conver;
import ctrl.fileFormat;
import java.util.Map;
import java.util.List;
import java.util.Arrays;
import java.util.HashMap;
/**
 * �ļ���ʽת�����ߣ�֧�ֱ༭word�ĵ��Ĺ���
 * 
 * @author Wenzhou
 *
 */
public class transfer
{
	// �ɴ��ļ�����
	private static final List<String> OPEN = Arrays.asList("ppt", "xls", "html", "xml", "pdf", "doc");
	// �����ļ���ʽ��ת������
	private static final List<String> DOC2 = Arrays.asList("doc", "docx", "pdf", "html", "xml", "rtf", "txt");
	private static final List<String> PPT2 = Arrays.asList("ppt", "pptx", "pdf", "bmp", "jpg", "rtf");
	private static final List<String> XLS2 = Arrays.asList("xls", "xlsx", "pdf", "html", "csv", "xml", "txt");
	private static final List<String> HTM2 = Arrays.asList("doc", "docx", "pdf", "xml", "txt");
	private static final List<String> XML2 = Arrays.asList("doc", "docx", "pdf", "html", "txt");
	private static final List<String> PDF2 = Arrays.asList("doc", "docx", "html", "xml", "txt");

	private JFrame frame;
	private JTextField textField; // ���ļ�
	private JTextField textField_1; // ����ļ�
	private JFileChooser open_chooser; // ѡ���ļ�
	private JFileChooser save_chooser;
	private String readPath; // �ļ�ԭʼ·��
	private String writePath; // �ļ�Ŀ��·��
	private fileFormat ff = null; // ��ȡ�ļ��ĸ�ʽ֮��
	private Map<String, List<String>> mapTo; // ��¼�ļ�ת��֮��ӳ��
	private editer edt = null;

	/**
	 * Launch the application.
	 */
//	public static void main(String[] args)
//	{
//		EventQueue.invokeLater(new Runnable()
//		{
//			public void run()
//			{
//				try
//				{
//					transfer window = new transfer();
//					window.frame.setVisible(true);
//				} catch (Exception e)
//				{
//					e.printStackTrace();
//				}
//			}
//		});
		
//		editer wordwin = new editer("D:\\Documents\\test.docx", "");
////		wordwin.setContent();
//		wordwin.setVisible(true);
//	}

	/**
	 * Create the application.
	 */
	public transfer()
	{
		initialize();
		transfer_view();
	}

	public void initialize()
	{
		// TODO Auto-generated method stub
		this.ff = new fileFormat();
		this.mapTo = new HashMap<String, List<String>>();
		this.mapTo.put("doc", DOC2);
		this.mapTo.put("pdf", PDF2);
		this.mapTo.put("xls", XLS2);
		this.mapTo.put("html", HTM2);
		this.mapTo.put("ppt", PPT2);
		this.mapTo.put("xml", XML2);
		this.mapTo.put("docx", DOC2);
		this.mapTo.put("xlsx", XLS2);
		this.mapTo.put("pptx", PPT2);
	}

	/**
	 * set the contents of the frame.
	 */
	private void transfer_view()
	{
		frame = new JFrame();
		frame.setBounds(100, 100, 545, 277);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(null);

		JLabel label = new JLabel("\u6587\u4EF6\uFF1A");
		label.setFont(new Font("����", Font.BOLD, 18));
		label.setBounds(26, 68, 57, 25);
		frame.getContentPane().add(label);

		JLabel lblNewLabel = new JLabel("\u4FDD\u5B58\u76EE\u5F55\uFF1A");
		lblNewLabel.setFont(new Font("����", Font.BOLD, 18));
		lblNewLabel.setBounds(10, 119, 95, 25);
		frame.getContentPane().add(lblNewLabel);

		textField = new JTextField();
		textField.setBounds(105, 68, 299, 25);
		frame.getContentPane().add(textField);
		textField.setColumns(10);

		textField_1 = new JTextField();
		textField_1.setBounds(105, 121, 299, 25);
		frame.getContentPane().add(textField_1);
		textField_1.setColumns(10);

		open_chooser = new JFileChooser();
		save_chooser = new JFileChooser();
		open_chooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);// ����ѡ��ģʽ���ȿ���ѡ���ļ��ֿ���ѡ���ļ���
		save_chooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
		
		// �� ��ť
		JButton open_file = new JButton("\u6253\u5F00");
		open_file.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
				formatFilter filter = new formatFilter(OPEN);
				filter.addFilter(open_chooser);
				int index = open_chooser.showOpenDialog(null);
				open_chooser.setDialogType(JFileChooser.OPEN_DIALOG);
				open_chooser.setMultiSelectionEnabled(false);
//				open_chooser.setAcceptAllFileFilterUsed(false);
				if (index == JFileChooser.APPROVE_OPTION)
				{
					// �ѻ�ȡ�����ļ��ľ���·����ʾ���ı��༭����
					textField.setText(open_chooser.getSelectedFile().getAbsolutePath());
					readPath = textField.getText();
				}
				filter.removeFilter(open_chooser);
			}
		});
		open_file.setFont(new Font("����", Font.BOLD, 18));
		open_file.setBounds(432, 67, 87, 26);
		frame.getContentPane().add(open_file);

		// ��� ��ť
		JButton browse = new JButton("\u6D4F\u89C8");
		browse.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
				String []name = checkFormat(textField.getText());	//��ȡ�ļ�������չ��
				formatFilter filter = new formatFilter(mapTo.get(name[1]));	//��ȡ��ת���ĸ�ʽ
				filter.addFilter(save_chooser);	// Ϊ���������filter
				save_chooser.setSelectedFile(new File(name[0]));

				int index = save_chooser.showSaveDialog(null);
				save_chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				save_chooser.setDialogType(JFileChooser.SAVE_DIALOG);
				save_chooser.setMultiSelectionEnabled(false);
				save_chooser.setAcceptAllFileFilterUsed(false);
				if (index == JFileChooser.APPROVE_OPTION)
				{
					// ��ȡĿ���ļ�����չ��
					myFilter cur_mf = (myFilter)save_chooser.getFileFilter(); // ��ù���������չ��
					String ends = "."+cur_mf.getDescription();
					// �ѻ�ȡ�����ļ��ľ���·����ʾ���ı��༭����
					writePath = save_chooser.getSelectedFile().getAbsolutePath();
					if(writePath.contains(name[0]+ends)==false)//����û������Ƿ������չ��
					{
						if(writePath.contains(name[0])==false)
						{
							writePath = writePath+"\\"+name[0];
						}
						if(writePath.contains(ends)==false)
							writePath = writePath + ends;
					}
					textField_1.setText(writePath);
				}
				filter.removeFilter(save_chooser);
			}
		});
		browse.setFont(new Font("����", Font.BOLD, 18));
		browse.setBounds(432, 118, 87, 26);
		frame.getContentPane().add(browse);
		
		// ���Ϊ ��ť
		JButton save_as = new JButton("\u53E6\u5B58\u4E3A");
		save_as.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
				readPath = textField.getText();
				writePath = textField_1.getText();
//				JOptionPane.showMessageDialog(null, "Դ�ļ�������", "����", JOptionPane.ERROR_MESSAGE);
				try
				{
					conver cv = new conver(readPath, writePath);
					cv.save();		// ������Ŀ��·��
				} catch (Exception e1)
				{
					// TODO Auto-generated catch block
					e1.printStackTrace();
					JOptionPane.showMessageDialog(null, "�ļ�ת������", "����", JOptionPane.ERROR_MESSAGE);
				}
			}
		});
		save_as.setForeground(Color.BLACK);
		save_as.setFont(new Font("����", Font.BOLD, 18));
		save_as.setBounds(293, 180, 93, 34);
		frame.getContentPane().add(save_as);
		
		// �༭ ��ť
		JButton edit = new JButton("�༭");
		edit.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
				readPath = textField.getText();
				writePath = textField_1.getText();
				String[] name = checkFormat(readPath);
				if(name[1].equals("doc")||name[1].equals("docx"))
				{
					edt = new editer(readPath, writePath);
					edt.setVisible(true);
				}
				else
				{
					// ������ʽ�ļ��༭֧�� ���ڴ˴����
					JOptionPane.showMessageDialog(null, "�ݲ�֧�ֱ༭�����ļ�", "����", JOptionPane.ERROR_MESSAGE);
				}
			}
		});
		edit.setForeground(Color.BLACK);
		edit.setFont(new Font("����", Font.BOLD, 18));
		edit.setBounds(143, 180, 93, 34);
		frame.getContentPane().add(edit);
		
		frame.addWindowListener(new WindowAdapter()
		{
			public void windowClosing(WindowEvent e)
			{
				if(transfer.this.edt != null)
				{
					transfer.this.edt.exit();
					System.exit(0);
				}
			}
		});
	}
/*
	public int copyFile(String oldPath, String newPath)
	{
		try
		{
			int bytesum = 0;
			int byteread = 0;
			File oldfile = new File(oldPath);
			if (oldfile.exists())
			{ // �ļ�����ʱ
				InputStream inStream = new FileInputStream(oldPath); // ����ԭ�ļ�
				System.out.println(newPath);
				if (isExist(newPath))
				{
					FileOutputStream fs = new FileOutputStream(newPath);
					byte[] buffer = new byte[1444];
					while ((byteread = inStream.read(buffer)) != -1)
					{
						bytesum += byteread; // �ֽ��� �ļ���С
						System.out.println(bytesum);
						fs.write(buffer, 0, byteread);
					}
					inStream.close();
					fs.close();
					return 0;
				} else
				{
					return -2;
				}
			}
			return -1;
		} catch (Exception e)
		{
			System.out.println("���Ƶ����ļ���������");
			e.printStackTrace();
			return -2;
		}
	}

	public static boolean isExist(String filePath)
	{
		String paths[] = filePath.split("\\\\");
		String dir = paths[0];
		for (int i = 0; i < paths.length - 2; i++)
		{// ע��˴�ѭ���ĳ���
			try
			{
				dir = dir + "/" + paths[i + 1];
				File dirFile = new File(dir);
				if (!dirFile.exists())
				{
					dirFile.mkdir();
					System.out.println("����Ŀ¼Ϊ��" + dir);
				}
			} catch (Exception err)
			{
				System.err.println("ELS - Chart : �ļ��д��������쳣");
			}
		}
		File fp = new File(filePath);
		if (!fp.exists())
		{
			return true; // �ļ������ڣ�ִ�����ع���
		} else
		{
			return false; // �ļ����ڲ�������
		}
	}
*/
	
	public String[] checkFormat(String cpath)
	{
		ff.setPath(cpath);
		if (ff.getFormat() != 0)
		{
			return null;
		} 
		else
			return ff.name;
	}
}