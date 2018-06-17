package view;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Desktop;
import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.datatransfer.DataFlavor;
import java.awt.dnd.DnDConstants;
import java.awt.dnd.DropTarget;
import java.awt.dnd.DropTargetAdapter;
import java.awt.dnd.DropTargetDropEvent;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTree;
import javax.swing.event.TreeSelectionEvent;
import javax.swing.event.TreeSelectionListener;
import javax.swing.tree.DefaultMutableTreeNode;
import javax.swing.tree.DefaultTreeCellRenderer;
import javax.swing.tree.DefaultTreeModel;

import ctrl.conver;
import ctrl.fileFormat;
import javax.swing.UIManager;

public class big_view
{
	private static final List<String> OPEN = Arrays.asList("ppt", "xls", "html", "xml", "pdf", "doc");
	
	private static final List<String> DOC2 = Arrays.asList("doc", "docx", "pdf", "html", "xml", "rtf", "txt");
	private static final List<String> PPT2 = Arrays.asList("ppt", "pptx", "pdf", "bmp", "jpg", "rtf");
	private static final List<String> XLS2 = Arrays.asList("xls", "xlsx", "pdf", "html", "csv", "xml", "txt");

	
	private Map<String, List<String>> mapTo; // ��¼�ļ�ת��֮��ӳ��


	private JFileChooser open_chooser; // ѡ���ļ�
	private JFileChooser save_chooser;
	
	private List<String> open;	// Դ�ļ���ʽ
	private List<String> save;	// Ŀ���ļ���ʽ
	
	private String readPath;
	private String writePath; // �ļ�Ŀ��·��
	
	private fileFormat ff = null; // ��ȡ�ļ��ĸ�ʽ֮��
	
	private JFrame frame;
	/**
	 * Launch the application.
	 */
	public static void main(String[] args)
	{
		EventQueue.invokeLater(new Runnable()
		{
			public void run()
			{
				try
				{
					big_view window = new big_view();
					window.frame.setVisible(true);
				} catch (Exception e)
				{
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public big_view()
	{
		this.ff = new fileFormat();
		this.mapTo = new HashMap<String, List<String>>();
		this.mapTo.put("doc", DOC2);
		this.mapTo.put("xls", XLS2);
		this.mapTo.put("ppt", PPT2);
		this.mapTo.put("docx", DOC2);
		this.mapTo.put("xlsx", XLS2);
		this.mapTo.put("pptx", PPT2);
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize()
	{
		frame = new JFrame();
		frame.setTitle("\u6587\u4EF6\u8F6C\u6362\u5668");
		frame.setBounds(200, 30, 950, 654);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(new BorderLayout(0, 0));

		JPanel panel2 = new JPanel();
		panel2.setBackground(new Color(0,0,0,0));
		frame.getContentPane().add(panel2, BorderLayout.CENTER);
		panel2.setLayout(null);
		
		// �ĵ�ѡ����
		open_chooser = new JFileChooser();
		save_chooser = new JFileChooser();
//		open_chooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);// ����ѡ��ģʽ���ȿ���ѡ���ļ��ֿ���ѡ���ļ���
//		save_chooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
		
		// �°벿��
		JPanel panel = new JPanel();
		panel.setBounds(0, 531, 684, 85);
		panel2.add(panel);
		panel.setLayout(null);
		
		JButton open_file = new JButton("\u6DFB\u52A0\u6587\u4EF6");
		open_file.setFont(new Font("����", Font.PLAIN, 12));
		open_file.setBounds(27, 20, 138, 44);
		panel.add(open_file);
		
		JButton save_file = new JButton("\u4FDD\u5B58\u76EE\u5F55");
		save_file.setBounds(203, 20, 129, 44);
		panel.add(save_file);
		
		JButton transe = new JButton("\u5F00\u59CB\u8F6C\u6362");
		transe.setBackground(new Color(100, 149, 237));
		transe.setForeground(Color.BLACK);
		transe.setBounds(370, 20, 118, 44);
//		btnNewButton_2.setBackground(new Color(255,255,10, 0));
		panel.add(transe);
		
		JButton openSys = new JButton("\u6253\u5F00\u6587\u4EF6\u6240\u5728\u76EE\u5F55");
		openSys.setBounds(521, 20, 138, 44);
		panel.add(openSys);
		
		// �м䲿��
		JPanel panel_1 = new HomePanel();
		panel_1.setBounds(0, 0, 684, 532);
		panel2.add(panel_1);

		panel_1.setLayout(null);
		
		JLabel input = new JLabel("");
		input.setBounds(26, 406, 624, 55);
		panel_1.add(input);
		
		JLabel output = new JLabel("");
		output.setBounds(26, 471, 624, 51);
		panel_1.add(output);
		
		DropTarget dropTarget = new DropTarget(panel_1, DnDConstants.ACTION_COPY_OR_MOVE,
			new DropTargetAdapter()
	        {
	            @Override
	            public void drop(DropTargetDropEvent dtde)
	            {
	               try
	               {
	                  // ���������ļ���ʽ��֧��
	                  if (dtde.isDataFlavorSupported(DataFlavor.javaFileListFlavor))
	                  {
	                     // ������ק��������
	                     dtde.acceptDrop(DnDConstants.ACTION_COPY_OR_MOVE);
	                     @SuppressWarnings("unchecked")
	                     List<File> list = (List<File>) (dtde.getTransferable()
	                           .getTransferData(DataFlavor.javaFileListFlavor));
	                     for (File file : list)
	                     {
	                    	 
	                    	 big_view.this.readPath = file.getAbsolutePath();
	                    	 input.setText(big_view.this.readPath);
	                     }
	                     // ָʾ��ק���������
	                     dtde.dropComplete(true);
	                  }
	                  else
	                  {
	                     // �ܾ���ק��������
	                     dtde.rejectDrop();
	                  }
	               }
	               catch (Exception e)
	               {
	                  e.printStackTrace();
	               }
	            }
	         });
		

		// ��벿��
		JPanel panel1 = new JPanel();
		panel1.setPreferredSize(new Dimension(250, 0));
		frame.getContentPane().add(panel1, BorderLayout.WEST);
		
		DefaultTreeCellRenderer myRenderer = new DefaultTreeCellRenderer();
		myRenderer.setOpenIcon(new ImageIcon("res\\table_2.png"));
		myRenderer.setClosedIcon(new ImageIcon("res\\table_1.png"));
		myRenderer.setLeafIcon(new ImageIcon("res\\table_3.png"));
		myRenderer.setFont(new Font("����", Font.BOLD, 18));
		myRenderer.setBackground(new Color(0,0,0,0));
		// һ���ڵ�
	    DefaultMutableTreeNode office_root = new DefaultMutableTreeNode(new JTreeData("office�ļ�ת��"));
	    DefaultMutableTreeNode other_root = new DefaultMutableTreeNode(new JTreeData("�����ļ�ת��"));
        // �����ڵ�
	    DefaultMutableTreeNode word2 = new DefaultMutableTreeNode(new JTreeData("word�ĵ�ת��", "doc"));
	    DefaultMutableTreeNode ppt2 = new DefaultMutableTreeNode(new JTreeData("ppt�ĵ�ת��", "ppt"));
	    DefaultMutableTreeNode excel2 = new DefaultMutableTreeNode(new JTreeData("excel�ĵ�ת��", "xls"));
	    DefaultMutableTreeNode pdf2word = new DefaultMutableTreeNode(new JTreeData("pdfתword�ĵ�", "pdf", "doc"));
	    DefaultMutableTreeNode html2word = new DefaultMutableTreeNode(new JTreeData("htmlתword�ĵ�", "html", "doc"));
	    DefaultMutableTreeNode html2pdf = new DefaultMutableTreeNode(new JTreeData("htmlתpdf�ĵ�", "html", "pdf"));
	    // �����ڵ�
	    for (String list : DOC2)
	    	word2.add(new DefaultMutableTreeNode(new JTreeData("ת����"+list+"�ĵ�", "doc", list)));
	    for (String list : PPT2)
	    	ppt2.add(new DefaultMutableTreeNode(new JTreeData("ת����"+list+"�ĵ�", "ppt", list)));
	    for (String list : XLS2)
	    	excel2.add(new DefaultMutableTreeNode(new JTreeData("ת����"+list+"�ĵ�", "xls", list)));
	    
	    office_root.add(word2);
	    office_root.add(ppt2);
	    office_root.add(excel2);
	    
	    other_root.add(pdf2word);
	    other_root.add(html2pdf);
	    other_root.add(html2word);
	    
        DefaultTreeModel officeModel = new DefaultTreeModel(office_root);
        DefaultTreeModel otherModel = new DefaultTreeModel(other_root);

        // ͨ��add()�����������ڵ�֮��ĸ��ӹ�ϵ

		panel1.setLayout(null);
		JTree tree1 = new JTree(officeModel);
		tree1.setBounds(0,0,250, 327);
		tree1.setCellRenderer(myRenderer);
		panel1.add(tree1);
		
		JTree tree2 = new JTree(otherModel);
		tree2.setBounds(0,327,250, 289);
		tree2.setCellRenderer(myRenderer);
		panel1.add(tree2);
		// ����Ӽ�����
		tree1.addTreeSelectionListener(new TreeSelectionListener() {
            @Override
            public void valueChanged(TreeSelectionEvent e) {
                DefaultMutableTreeNode node = (DefaultMutableTreeNode) tree1
                        .getLastSelectedPathComponent();
 
                if (node == null)
                    return;
 
                Object object = node.getUserObject();
                JTreeData user = (JTreeData) object;
                if (user.getSource()!=null)
                {
                	big_view.this.open = Arrays.asList(user.getSource());
                }
                else
                	big_view.this.open = null;
                if (user.getTarget()!=null)
                {
                	big_view.this.save = Arrays.asList(user.getTarget());
                }
                else
                	big_view.this.save = null;
            }
        });
		tree2.addTreeSelectionListener(new TreeSelectionListener() {
            @Override
            public void valueChanged(TreeSelectionEvent e) {
                DefaultMutableTreeNode node = (DefaultMutableTreeNode) tree2
                        .getLastSelectedPathComponent();
 
                if (node == null)
                    return;
 
                Object object = node.getUserObject();
                JTreeData user = (JTreeData) object;
                if (user.getSource()!=null)
                {
                	big_view.this.open = Arrays.asList(user.getSource());
                }
                else
                	big_view.this.open = null;
                if (user.getTarget()!=null)
                {
                	big_view.this.save = Arrays.asList(user.getTarget());
                }
                else
                	big_view.this.save = null;
            }
        });
		// ��ť��Ӽ�����
		open_file.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
				formatFilter filter = null;
				if (big_view.this.open!=null)	// �����openҪ���ó�list
					filter = new formatFilter(big_view.this.open);
				else
					filter = new formatFilter(OPEN);
				filter.addFilter(open_chooser);
				int index = open_chooser.showOpenDialog(null);
				open_chooser.setDialogType(JFileChooser.OPEN_DIALOG);
				open_chooser.setMultiSelectionEnabled(false);
//				open_chooser.setAcceptAllFileFilterUsed(false);
				if (index == JFileChooser.APPROVE_OPTION)
				{
					// �ѻ�ȡ�����ļ��ľ���·����ʾ���ı�����
					readPath = open_chooser.getSelectedFile().getAbsolutePath();
					input.setText(readPath);
				}
				filter.removeFilter(open_chooser);
			}
		});
		
		save_file.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
				String []name = checkFormat(readPath);	//��ȡ�ļ�������չ��
				formatFilter filter = null;
				// ���û��ѡ��Ŀ���ʽ
				if (big_view.this.save!=null)
					filter = new formatFilter(big_view.this.save);
				else
					filter = new formatFilter(mapTo.get(name[1]));
				filter.addFilter(save_chooser);	// Ϊ���������filter
				save_chooser.setSelectedFile(new File(name[0]));

				int index = save_chooser.showSaveDialog(null);
				save_chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
				save_chooser.setDialogType(JFileChooser.SAVE_DIALOG);
				save_chooser.setMultiSelectionEnabled(false);
				save_chooser.setAcceptAllFileFilterUsed(false);
				if (index == JFileChooser.APPROVE_OPTION)
				{
					// ��ȡĿ���ļ�����չ��
					myFilter cur_mf = (myFilter)save_chooser.getFileFilter(); // ��ù���������չ��
					String ends = "."+cur_mf.getDescription();
					// �ѻ�ȡ�����ļ��ľ���·����ʾ���ı�����
					writePath = save_chooser.getSelectedFile().getAbsolutePath();
//					if(writePath.contains(name[0]+ends)==false)//����û������Ƿ������չ��
//					{
//						if(writePath.contains(name[0])==false)
//						{
//							writePath = writePath+"\\"+name[0];
//						}
//						if(writePath.contains(ends)==false)
//							writePath = writePath + ends;
//					}
					writePath = writePath + ends;
					output.setText(writePath);
				}
				filter.removeFilter(save_chooser);
			}
		});
		
		transe.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
//				JOptionPane.showMessageDialog(null, "Դ�ļ�������", "����", JOptionPane.ERROR_MESSAGE);
				try
				{
					conver cv = new conver(readPath, writePath);
					cv.save();		// ������Ŀ��·��
				} catch (Exception e1)
				{
					// TODO Auto-generated catch block
					e1.printStackTrace();
//					JOptionPane.showMessageDialog(null, "�ļ�ת������", "����", JOptionPane.ERROR_MESSAGE);
				}
			}
		});
		
		openSys.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
//				JOptionPane.showMessageDialog(null, "Դ�ļ�������", "����", JOptionPane.ERROR_MESSAGE);
				try
				{
					Desktop.getDesktop().open(new File(getSysPath(writePath)));  
				} catch (Exception e1)
				{
					// TODO Auto-generated catch block
					e1.printStackTrace();
//					JOptionPane.showMessageDialog(null, "�ļ�ת������", "����", JOptionPane.ERROR_MESSAGE);
				}
			}
		});
	}
	
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
	
	public String getSysPath(String cpath)
	{
		ff.setPath(cpath);
		if (ff.getFormat() != 0)
		{
			return null;
		} 
		else
			return ff.sys_path;
	}
}

