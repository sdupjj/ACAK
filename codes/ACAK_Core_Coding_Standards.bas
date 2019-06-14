Attribute VB_Name = "ACAK_Core_Coding_Standards"
'ACAK Core 代码编写规则

'所有ACAK中的sheet 名字全部以 core_ 开头
'所有参考的网页资料 归档到 ACAK_Thanks_List，格式: 一行标题 一行网址
'所有的代码 全部归到 func_core_010, 020, 030, 040, 990  中 （sheet代码除外）
'每个sub/function 开头 都要标注以下内容：
'''
'1功能：
'1  用数字来获取英语字母
'1当前版本：
'1  1.1
'1历史版本修订内容：
'1  1.1 >>> 新版本
'1  1.0 >>> 原始版本
'2定义
'2赋值
'2代码
'''
'func_core_010 中存放 One 引擎的核心代码，每个sub/function 以 c_ 开头
'func_core_020 中存放 interface的代码，可以被 core_if 调用，每个sub 以 if_ 开头
'func_core_030 中存放 supports的代码, 仅可以被别的代码调用，以及 按钮 加载宏 调用，每个sub/function 以 cs_ 开头
'func_core_040 中存放 actions的代码，可以被M02 调用， “core_engine_setup”与 "core_ACAK_setup" 调用， 每个sub 以 a_开头
'func_core_990 中存放 debug的代码，目标是为了core_debug 调用使用 ，每个sub/function 以 cs_ 开头
'sub/function 中每个变量都要 以cv_开头
