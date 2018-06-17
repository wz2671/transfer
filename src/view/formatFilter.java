package view;

import java.io.File;
import java.io.FilenameFilter;
import java.util.List;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileNameExtensionFilter;
/**
 * �ļ����͹�����
 * @author Wenzhou
 *
 */
public class formatFilter 
{
	private int length;
	private String[] filters=new String[100];
	private myFilter []mf; 

	public formatFilter(){}
	/**
	 * ��ʼ��������
	 * @param str	�������ļ���չ��
	 */
	public formatFilter(List<String> str)
	{
		this.filters=(String[]) str.toArray();
		this.length=str.size();
		this.mf = new myFilter[this.length];
		for(int i=0; i<this.length; i++)
		{
			switch(this.filters[i])
			{
			case "doc":
				this.mf[i] = new myFilter(new String[]{"doc","docx"},"doc");
				break;
			case "xls":
				this.mf[i] = new myFilter(new String[]{"xls","xlsx"},"xls");
				break;
			case "ppt":
				this.mf[i] = new myFilter(new String[]{"ppt","pptx"},"ppt");
				break;
			default:
				this.mf[i] = new myFilter(this.filters[i], this.filters[i]);
			}
		}
	}
	/**
	 * ����ļ�������
	 * @param chooser	Ҫ��ӵ��ļ�ѡ����
	 */
	public void addFilter(JFileChooser chooser)
	{
		for(int i=0; i<this.length; i++)
		{
			chooser.setFileFilter(this.mf[i]);
		}
	}
	/**
	 * �Ƴ��ļ�������
	 * @param chooser
	 */
	public void removeFilter(JFileChooser chooser)
	{
		for(int i=0; i<this.length; i++)
		{
			chooser.removeChoosableFileFilter(mf[i]);
		}
	}
}

