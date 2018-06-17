package ctrl;

import java.util.Arrays;

/**
 * ��·�����ļ�������չ�����зָ�
 * @author Wenzhou
 *
 */
public class fileFormat 
{
	//���ڽ���Ŀ���ļ���ʽ
	public String msg = null;
	public String target_path = null;
	//��ǰ��Ҫת�����ļ�·��
	public String current_path = null;
	//ϵͳ·��
	public String sys_path = null;
	//Ĭ�ϵ��ļ����͸�ʽ��׺
	public String[] name = null;
	
	public fileFormat()
	{
		// ��ʱ��ʲô������
	}
	
	//�����ֱ�Ϊԭʼ�ļ�·�������ļ���
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
		// ����Խ��Ļ���������ļ���������
		try
		{	
			name = (cpath_sp[cpath_sp.length -1]).split("\\.");
			cpath_sp[cpath_sp.length-1] = name[0];
			//���д洢�Ż�û���ļ���չ����·��
			target_path = String.join("\\", cpath_sp);
//			type = cpath_sp[cpath_sp.length-1];
			String[] csys_path = Arrays.copyOfRange(cpath_sp, 0, cpath_sp.length-1);
			this.sys_path = String.join("\\", csys_path);
		}
		catch(Exception e)
		{
			msg = "�ļ���ʽ����";
			System.out.println(msg);
			e.printStackTrace();
			return -1;
		}
		return 0;
	}
}
