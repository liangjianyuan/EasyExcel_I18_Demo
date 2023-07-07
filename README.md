# EasyExcel_I18_Demo
EasyExcel国际化demo，对导入、导出表头做国际化

# 项目结构
调用Deamo 在测试类 MessagesTest中，包含导入导出类
核心功能类在ExcelUtil

# 实现思路
在easyexcel基础上对导入、导出对头做处理，导出时向writer对象注册handler进行处理；导入写方法构建新的头达到和注解中写的一样来读取文件，demo中采用sprng的国际化来存储不同语言的映射；
