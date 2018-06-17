package view;

import java.io.File;
import java.io.FilenameFilter;
import java.util.List;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileNameExtensionFilter;
/**
 * 文件类型过滤器
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
	 * 初始化过滤器
	 * @param str	待过滤文件扩展名
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
	 * 添加文件过滤器
	 * @param chooser	要添加的文件选择器
	 */
	public void addFilter(JFileChooser chooser)
	{
		for(int i=0; i<this.length; i++)
		{
			chooser.setFileFilter(this.mf[i]);
		}
	}
	/**
	 * 移除文件过滤器
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

