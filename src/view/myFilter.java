package view;

import java.io.File;
import java.io.FilenameFilter;
import java.util.List;

import javax.swing.filechooser.FileFilter;
/**
 * 真·文件类型过滤器
 * @author Wenzhou
 *
 */
public class myFilter extends FileFilter
{
		private int length;
		private String[] filters=new String[100];
		private String desc;

		public myFilter(){}

		public myFilter(String str)
		{
			this.filters[length]=str;
			length++;
		}
		
		public myFilter(String str,String desc)
		{
			this.filters[length]=str;
			this.desc=desc;
			length++;
		}

		public myFilter(String str[],String desc)
		{
			this.filters=str;
			this.desc=desc;
			this.length=str.length;
		}
		
		public boolean accept(File f) 
		{
			String tmp=f.getName().toLowerCase();
			//如果小于0 显示所有文件
			if(length==0)
			{
				return true; 
			}
			//显示文件夹
			if(f.isDirectory())
			{
				return true; 
			}
			//循环过滤文件过滤
			for(int i=0;i<length;i++)
			{
				if(tmp.endsWith(this.filters[i]))
				{
					return true;
				}
			}
			return false;
		}
	
		/**
		* 
		* @param str 过滤器名称 例如:".doc"
		*/
		public void addFilter(String str)
		{
			this.filters[length]=str;
			this.length++;
		}
		/**
		* 
		* @param str 过滤器名称 例如:".doc"
		* @param desc @param desc 此过滤器的描述。例如："word文档" 
		*/
		public void addFilter(String str,String desc)
		{
			this.filters[length]=str;
			this.desc=desc;
			length++;
		}
	
		/** 
		* @param str 传递多个 过滤器名称 例如:{".doc",".docx"}
		* @param desc 此过滤器的描述。例如："压缩文件" 
		*/
		public void addFilter(List<String> str,String desc)
		{
			this.filters=(String[]) str.toArray();
			this.desc=desc;
			this.length=str.size();
			for(int i=0; i<this.length; i++)
			{
				System.out.println(this.filters);
			}
		}
	
		/**
		* 
		* @param desc 此过滤器的描述。 
		*/
		public void setDesc(String desc)
		{
			this.desc=desc;
		} 
	
		public String getDescription() 
		{
			return desc;
		} 
}
