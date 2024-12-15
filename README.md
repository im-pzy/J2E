# J2E
将Json Array转换为Excel表格的工具


**nuitka打包命令：**
```
nuitka --standalone --onefile --windows-console-mode=disable --enable-plugin=pyqt5 --include-package=xlsxwriter --windows-icon-from-ico=".\icons\J2E.ico" --include-data-files=".\icons\*=icons/" J2E.py
```

# 有待改进之处

1. json预览功能
2. excel转换为json功能
3. 表格用QAbstractTableModel配合QTableView来实现，表格的变动会同步到数据，同时这样还有利于内存优化
4. 表格可以按行、按列、按选区复制
5. 支持单引号
6. 支持注释
