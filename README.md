# Excel-VBA-Tools

## EEPROMXlsx2Bin
自己工作用遇到的重复工作, 使用VBA脚本实现些许自动化
### 背景
高速连接器模块在出厂前均需要烧录EEPROM, 其内容会根据不同客户的需求产生差异, 所以会采用套用模版的方法, 即Excel xlsx文件, 编辑后需要将其内容导出为TXT或者BIN文件格式.


### 使用方法
- 将xlam文件存储到本地后, 打开Excel, 在开发工具中选择加载宏, 同时选中该xlam文件中提供的脚本
- 使用自定义工具栏绑定该脚本, 即点击工具栏图标即可实现调用此脚本
- 在模版中使用此工具即可在模版文件所在目录中生成到处的TXT和BIN文件.

### 环境要求
- Office Excel 版本不低于2016.
- Office Excel 需要打开宏支持.

## SalaryAutoGen
财务可能用的到的工资条生成器

## SNAutoChecker
产线出货SN扫码校验器 
