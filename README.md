# excel_analysis
### 中文
Python自动化处理Excel，用于提取表格的项
支持正则表达式，与或非语法逻辑，目前仅支持字符串级别的数值处理

使用规则：
excel_analysis filepath [exportpath] --sheet 0 1 --rule "A1:D:(aa&bb)|(cc)"  "B1:S:!([0-9]+)|([a-z]{4})" ... --sheet n  --rule "A1:D:(aa&bb)|(cc)"  "B1:S:([0-9]+)|([a-z]{4})"

filepath
   操作的excel表格对象
   
exportpath
   可选，保存的文件路径，可以包含文件名，没有则是result.xlsx
   
 --sheet ...
     n  操作对象是第n个sheet,并将后面的规则应用到第n个sheet对象中，可以同时应用任意个sheet
     
 --rule ...
     A4 表示对第A列第4行及以下做处理，往下所有规则应用
     D/S 表示Delete/Save, 对按规则匹配到的行执行删除或是保留操作（执行删除时，保留剩下行；执行保留时，删除剩余行）
     正则表达式定义的规则   & ｜ ！ 表示 与 或 非  建议!用括号括起，防止歧义 
