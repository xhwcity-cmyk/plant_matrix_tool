# PlantMatrix 物种数据整理工具

# 主要功能
## Excel专用处理：
直接读取Excel文件，不经过文本转换
使用pandas的Excel读取功能
保持原始数据格式和结构

## 智能表格识别：
自动查找所有包含"物种"的表头
识别每个独立表格的边界
正确处理表格间的空白区域

## 数据提取和清理：
准确提取物种名称和对应数值
自动处理缺失值和空单元格
数值类型转换和验证

## 调试和分析功能：
详细的文件结构分析
处理过程日志输出
可视化调试信息

## 使用方法
点击"选择Excel文件并处理"按钮选择您的Excel文件
程序会自动识别所有表格并合并
查看处理结果和输出文件
如果遇到问题，可以使用"分析Excel文件结构"按钮来查看文件详细结构
这个工具应该能够正确处理您提供的Excel格式的植物样方数据。如果仍有问题，请使用调试功能分析文件结构，这样我可以更准确地了解数据格式并提供进一步帮助。
### 表格输入示例
示例表格已经过修改，无任何实质性内容。
<img width="209" height="1065" alt="image" src="https://github.com/user-attachments/assets/86f8ee75-632f-4028-891d-2a802dc98c0c" />

<img width="690" height="449" alt="image" src="https://github.com/user-attachments/assets/16541b5d-54dc-4f47-8898-e9aa3550beec" />

## 下载
前往 [Releases](https://github.com/你的用户名/你的仓库/releases) 页面下载最新版本。

## 构建说明
本项目使用 GitHub Actions 自动构建多平台版本。
