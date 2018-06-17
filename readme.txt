1. 将jacob里的jar包导入到项目里， project->properties->java build path->add external jars->选中jacob.jar包
2. 如果是32位系统 将jacob-1.18-x86.dll 文件复制到下面目录下，如果是64位操作系统将jacob-1.18-x64.dll
	java\jdk1.8.0_45\jre\bin 
	不行就放在c:/windows目录下的 Sytem32/SysWOW64 对应文件夹下
	可参考博客：https://blog.csdn.net/zmx729618/article/details/61196749
3. 程序的入口类为view包中transfer类
4. 实现的功能： (1). 常用office文件类型转换(部分格式转换尚未测试)
		(2). 对word文档的简单编辑(接口函数已给出，其余功能可自行添加)