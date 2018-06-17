package ctrl;

import java.util.Arrays;

/**
 * 对路径，文件名，扩展名进行分割
 * @author Wenzhou
 *
 */
public class fileFormat 
{
	//用于接收目标文件格式
	public String msg = null;
	public String target_path = null;
	//当前需要转化的文件路径
	public String current_path = null;
	//系统路径
	public String sys_path = null;
	//默认的文件名和格式后缀
	public String[] name = null;
	
	public fileFormat()
	{
		// 暂时先什么都不做
	}
	
	//参数分别为原始文件路径包括文件名
	public fileFormat(String cpath)
	{
		current_path = cpath;
	}
	
	public void setPath(String cpath)
	{
		this.current_path = cpath;
	}
	
	public int getFormat()
	{
		String[] cpath_sp= null;
		try
		{
			cpath_sp = this.current_path.split("\\\\");
		}
		catch (Exception e)
		{
			e.printStackTrace();
			return -1;
		}
//		String type = null;
		// 数组越界的话，输入的文件名有问题
		try
		{	
			name = (cpath_sp[cpath_sp.length -1]).split("\\.");
			cpath_sp[cpath_sp.length-1] = name[0];
			//其中存储着还没有文件扩展名的路径
			target_path = String.join("\\", cpath_sp);
//			type = cpath_sp[cpath_sp.length-1];
			String[] csys_path = Arrays.copyOfRange(cpath_sp, 0, cpath_sp.length-1);
			this.sys_path = String.join("\\", csys_path);
		}
		catch(Exception e)
		{
			msg = "文件格式有误";
			System.out.println(msg);
			e.printStackTrace();
			return -1;
		}
		return 0;
	}
}
