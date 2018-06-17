package ctrl;

public abstract class bean
{
	public bean(){}
	/**
	 * @param s0 源文件路径
	 * @param s1 源文件扩展名
	 * @param s2 目标文件路径
	 * @param s3 目标文件扩展名
	 */
	public abstract void conversion(String s0, String s1, String s2, String s3);
	
	public abstract void close();
}
