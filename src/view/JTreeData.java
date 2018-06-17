package view;

import javax.swing.ImageIcon;
import javax.swing.JComponent;

public class JTreeData
{
	private String strNode;
	private String source;		// Դ�ļ���ʽ
	private String target;		// Ŀ���ļ���ʽ
	
	public JTreeData(String str)
	{
		strNode = str;
		this.source = null;
		this.target = null;
	}
	
	public JTreeData(String str, String sou)
	{
		strNode = str;
		this.source = sou;
		this.target = null;
	}

	public JTreeData(String str, String sou, String tar)
	{
		this.strNode = str;
		this.source = sou;
		this.target = tar;
	}

	public String getString()
	{
		return strNode;
	}

	public void setString(String strNode)
	{
		this.strNode = strNode;
	}

	public String getSource()
	{
		return source;
	}

	public String getTarget()
	{
		return target;
	}
	
	public String toString()
	{
		return this.strNode;
	}
	
}
