package ctrl;

import javax.swing.JOptionPane;

/**
 * ��ɸ�ʽת���Ĺ���,������gui����
 * @author Wenzhou
 *
 */
public class conver
{
	private String readPath = null;
	private String writePath = null;
	private fileFormat ff = new fileFormat();
	private String target_format = null;		//Ŀ���ļ���ʽ
	private String current_format = null;		//Դ�ļ���ʽ
	private bean b;
	private boolean done = true;
	
	/**
	 * ��ʼ��
	 * @param rp  Դ�ļ���·��
	 * @param wp  Ŀ���ļ�·��
	 */
	public conver(String rp, String wp) throws Exception
	{
		this.readPath = rp;
		this.writePath = wp;
		ff.setPath(readPath);
		if(ff.getFormat()==0)
		{
			this.current_format = ff.name[1];
		}
		else
		{
			this.done = false;
			JOptionPane.showMessageDialog(null, "Դ�ļ���ʽ����", "����", JOptionPane.ERROR_MESSAGE);
		}
		
		if(this.current_format.equals("ppt")||this.current_format.equals("pptx"))
		{
			this.b = new pptBean(this.readPath, true);	//�ƺ���������
		}
		else if(this.current_format.equals("xls")||this.current_format.equals("xlsx"))
		{
			this.b = new excelBean();
		}
		else
		{
			this.b = new wordBean();
		}
		
		ff.setPath(writePath);
		if(ff.getFormat() == 0)
		{
			this.target_format = ff.name[1];
		}
		else
		{
			this.done = false;
			JOptionPane.showMessageDialog(null, "Ŀ���ļ���ʽ����", "����", JOptionPane.ERROR_MESSAGE);
		}
	}
	
	public void save()
	{
		if(this.done)
		{
			try
			{
				this.b.conversion(this.readPath, this.current_format, this.writePath, this.target_format);
				this.b.close();
				JOptionPane.showMessageDialog(null, "�ļ�����ɹ�", "֪ͨ", JOptionPane.WARNING_MESSAGE);
			}
			catch(Exception e)
			{				
				e.printStackTrace();
				JOptionPane.showMessageDialog(null, "�ļ�����ʧ��", "֪ͨ", JOptionPane.ERROR_MESSAGE);
				this.b.close();
			}
		}
	}
}
