package view;

import java.io.File;
import java.io.FilenameFilter;
import java.util.List;

import javax.swing.filechooser.FileFilter;
/**
 * �桤�ļ����͹�����
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
			//���С��0 ��ʾ�����ļ�
			if(length==0)
			{
				return true; 
			}
			//��ʾ�ļ���
			if(f.isDirectory())
			{
				return true; 
			}
			//ѭ�������ļ�����
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
		* @param str ���������� ����:".doc"
		*/
		public void addFilter(String str)
		{
			this.filters[length]=str;
			this.length++;
		}
		/**
		* 
		* @param str ���������� ����:".doc"
		* @param desc @param desc �˹����������������磺"word�ĵ�" 
		*/
		public void addFilter(String str,String desc)
		{
			this.filters[length]=str;
			this.desc=desc;
			length++;
		}
	
		/** 
		* @param str ���ݶ�� ���������� ����:{".doc",".docx"}
		* @param desc �˹����������������磺"ѹ���ļ�" 
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
		* @param desc �˹������������� 
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
