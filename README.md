Macro Office Automation Toolkit
Project Introduction
This project provides a statistics verification system for medical papers based on Office macro languages, aiming to address the issues of low efficiency, poor accuracy, and high statistical knowledge requirements for editors in the current verification process of medical journals. By utilizing Office macro languages (including Microsoft VBA macros or WPS Office JavaScript macros), this system can automatically verify common statistical analysis methods in medical journal papers within word processing software (such as Microsoft Word or WPS Office). The statistical methods that can be implemented include chi-square tests (2×2 and 2×3), t-tests, analysis of variance, logistic regression, Cox regression, and other commonly used medical statistics methods.
Function List
Users can select the data to be verified using the mouse to trigger the macro command. The macro command will automatically complete the statistical analysis verification and display the results directly in the word processing software interface (comment), achieving one-click automatic detection without the need for additional external software or tools.
Usage
Environment Dependencies
Microsoft Office 2016+ (VBA macros) for Windows systems (macOS/Linux are not supported temporarily)
WPS Office 12+ (JS macros) for Windows, Linux, and other operating systems
Installation Steps
Clone the repository to your local machine:
bash
git clone https://github.com/zooqiang/Macro.git



# Macro 办公自动化工具库  
## 项目简介  
本项目是提供一种基于Office宏语言的医学论文统计数据检测系统，旨在解决现有技术中医学期刊编辑对论文统计数据核查效率低、准确性差、对编辑人员统计知识要求高的问题。
通过利用Office宏语言（包括Microsoft VBA宏或WPS Office JavaScript宏），在文字处理软件（如Microsoft Word或WPS Office）中实现对医学期刊论文中常见统计分析方法的自动验证。本项目可实现的统计分析方法包括但不限于卡方检验（2×2和2×3）、t检验、方差分析、logistic回归和cox回归等常用医学统计方法。
。  

## 功能列表  
- 用户可通过鼠标选取待验证的数据，触发宏命令执行，宏命令自动完成统计分析验证，并将验证结果直接显示于文字处理软件界面中（批注），实现一键式自动检测，无需借助额外的外部软件或工具。

## 使用方法  
### 环境依赖  
- Microsoft Office 2016+（VBA 宏）  基于Windows 系统（暂不支持 macOS/Linux）  
- WPS Office 12+ （JS宏） 基于Windows 系统、Linux等多操作系统  


### 安装步骤  
1. 克隆仓库到本地：  
   ```bash
   git clone https://github.com/zooqiang/Macro.git
