package ctrl;

import javax.swing.JOptionPane;

/**
 * 完成格式转换的功能,不包含gui界面
 * @author Wenzhou
 *
 */
public class conver
{
	private String readPath = null;
	private String writePath = null;
	private fileFormat ff = new fileFormat();
	private String target_format = null;		//目标文件格式
	private String current_format = null;		//源文件格式
	private bean b;
	private boolean done = true;
	
	/**
	 * 初始化
	 * @param rp  源文件的路径
	 * @param wp  目标文件路径
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
			JOptionPane.showMessageDialog(null, "源文件格式有误", "警告", JOptionPane.ERROR_MESSAGE);
		}
		
		if(this.current_format.equals("ppt")||this.current_format.equals("pptx"))
		{
			this.b = new pptBean(this.readPath, true);	//似乎不给隐藏
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
			JOptionPane.showMessageDialog(null, "目标文件格式有误", "警告", JOptionPane.ERROR_MESSAGE);
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
				JOptionPane.showMessageDialog(null, "文件保存成功", "通知", JOptionPane.WARNING_MESSAGE);
			}
			catch(Exception e)
			{				
				e.printStackTrace();
				JOptionPane.showMessageDialog(null, "文件保存失败", "通知", JOptionPane.ERROR_MESSAGE);
				this.b.close();
			}
		}
	}
}
