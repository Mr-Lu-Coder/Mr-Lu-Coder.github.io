模式1：
行模式：将主表和附表中的行合并（必须列主键（第一行）相等）
功能1：将主表不存在的行加到主表中
功能2：将两表表同时存在的行相加（每行执行前 先询问）
功能3：将两表将两表所有同时存在的行相加（只询问第一次）

模式2：
列模式：将主表和附表中的列合并（必须行主键（第一列）相等）
功能1：将主表不存在的列加到主表中
功能2：将两表表同时存在的列相加（每行执行前 先询问）"
功能3：将两表将两表所有同时存在的列相加（只询问第一次）
模式3：
将一个目录下所有的EXcel合并
生成demo_col.xls文件
默认将所有行主关键字相同对应的那一列数字相加
如果列表表头含有关键字'姓', '名', '次', '号', '分', '卡'
则不相加
在命令行下只需将文件夹拖入黑框或直接输入