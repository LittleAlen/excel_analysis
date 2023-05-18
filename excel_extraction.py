# 使用规则：
# python excel_extraction filepath [exportpath] --sheet 0 1 --rule "A1:D:(aa&bb)|(cc)"  "B1:S:!([0-9]+)|([a-z]{4})" ... --sheet n  --rule "A1:D:(aa&bb)|(cc)"  "B1:S:([0-9]+)|([a-z]{4})"
# filepath
#   操作的excel表格对象
# exportpath
#   可选，保存的文件路径，可以包含文件名，没有则是result.xlsx
# --sheet ...
#     n  操作对象是第n个sheet,并将后面的规则应用到第n个sheet对象中，可以同时应用任意个sheet
# --rule ...
#     A4 表示对第A列第4行及以下做处理，往下所有规则应用
#     D/S 表示Delete/Save, 对按规则匹配到的行执行删除或是保留操作（执行删除时，保留剩下行；执行保留时，删除剩余行）
#     正则表达式定义的规则   & ｜ ！ 表示 与 或 非  建议!用括号括起，防止歧义 
#规则测试样例：
#(((!(((a))))))    !((Aa)&(B))|!C   
import xlwings as xw
import re
import sys

class Args:
    def __init__(self) -> None:
        self.filepath=""
        self.exportpath=""
        self.sheets=[]
        self.rules=[]
        args = sys.argv
       
        self.filepath=args[1]
        index=2
        if args[2]!="--sheet":
            self.exportpath=args[2]
            index+=1
        else:
            self.exportpath="./result.xlsx"

        while(index<len(args)):
            if args[index]=="--sheet":
                index+=1
                sheet=[]
                while(index <len(args) and args[index]!="--rule" ):
                    sheet.append(int(args[index]))
                    index+=1
                self.sheets.append(sheet)
            elif args[index]=="--rule":
                index+=1
                rule=[]
                while(index <len(args) and args[index]!="--sheet" ):
                    rule.append(Rule(args[index]))
                    index+=1
                self.rules.append(rule)
        
            
                
        


class Rule:
    # 使用解析函数解析表达式，然后通过与或函数来执行判断
    def __init__(self,rule:str) -> None:
        self.rule=rule
        match = re.match(r"^([A-Z]+)(\d+):([DS]):(.*)$", rule)
        self.column=match.group(1)
        self.start_row=match.group(2)
        self.action=match.group(3)
        self.pattern=match.group(4)
    def extract(self,rule):  #需要注意(a)|(b)情况  整体来看
        if rule[0]=="(":
            brackets=1
        else:
            return rule
        left=0
        right=len(rule)-1
        index=0
        while(index<len(rule)):
            if rule[index]=="(":
                brackets+=1
            elif rule[index]==")":
                brackets-=1
                if brackets ==0:
                    break
            index+=1
            
        if index==right:
            return self.extract(rule[1:-1])
        else:
            return rule
            
        
    def parse(self,value:str,rule=""):#先简化再处理
        if rule=="":
            rule=self.pattern
        rule=self.extract(rule)#剥离括号
        left=0
        right=len(rule)
        brackets=0
        split=-1
        for i in range(left,right):
            if  rule[i]== "(":
                brackets+=1
            elif rule[i]== ")":
                brackets-=1
            elif brackets==0:
                if rule[i]=="&" or rule[i]=="|":
                    split=i
        if split==-1: #没有与或逻辑
            if rule[left]=="!":
                return not self.parse(value,rule[left+1:])
            else:
                if re.fullmatch(rule,value)==None:
                    return False
                else:
                    return True
        else:
            if rule[split]=="&":
                return self.parse(value,rule[left:split]) and self.parse(value,rule[split+1:])
            elif rule[split]=="|":
                return self.parse(value,rule[left:split]) or self.parse(value,rule[split+1:])
    def print_info(self):
         print(f"INFO: 操作对象：{self.column} 起始行：{self.start_row} 规则：{ '保留' if self.action=='S' else '删除'} 模式：{self.pattern}")
               
class Sheet:
    def __init__(self,sheet):
        self.sheet=sheet
        num_rows, num_cols = sheet.range('A1').current_region.shape
        self.row=num_rows
        self.col=num_cols
    def apply(self,rules,target_sheet):
        print(f"INFO：----------规则应用----------")
        self.print_info()
        target_sheet.range("A1").value=self.sheet.used_range.value
        last_row=self.row
        target_terms=[1]*last_row

        for i in range(len(rules)):
            rules[i].print_info()
            values=self.sheet.range(f"{rules[i].column}1:{rules[i].column}{last_row}").options(numbers=int).value
            for j in range(int(rules[i].start_row)-1,last_row):
                if target_terms[j]==0:
                    continue
                if rules[i].action=="S":
                    # print(f"{values[j]}  {rules[i].parse(str(values[j]))}")
                    if not rules[i].parse(str(values[j])):
                        target_terms[j]=0
                elif rules[i].action=="D":
                    if rules[i].parse(str(values[j])):
                        target_terms[j]=0

        nums=0
        for i in range(len(target_terms)):
            if target_terms[i]==0:
                target_sheet.range(f"{i+1-nums}:{i+1-nums}").delete()
                nums+=1
        print(f"INFO：------------结束-------------")
        return
    def print_info(self):
        print(f"INFO: sheet表 {self.sheet.name} 行数 {self.row}  列数 {self.col}") 

class Book:
    def __init__(self,filepath,sheets_index=[[0]]):
        self.app = xw.App(visible=True, add_book=False)
        self.sheets=[]
        # print(filepath)
        self.book = self.app.books.open(filepath) # 打开Excel文件
        # print(self.book.name)
        # print(sheets_index)  #debug
        # print(type(self.book))
        for i in range(len(sheets_index)):
            sheet=[]
            for j in range(len(sheets_index[i])):
                sheet.append(Sheet(self.book.sheets[sheets_index[i][j]]))
            self.sheets.append(sheet)
    
    def apply_rules(self,exportpath,sheets_rule):
        new_book=self.app.books.add()
        target_sheet=""
        for i in range(len(self.sheets)):
            for j in range(len(self.sheets[i])):
                if target_sheet=="":
                    target_sheet=new_book.sheets[0]
                else:
                    target_sheet=new_book.sheets.add()
                target_sheet.name=self.sheets[i][j].sheet.name
                self.sheets[i][j].apply(sheets_rule[i],target_sheet)
        
        new_book.save(exportpath)
        new_book.close()
    def close(self):
        self.book.close()

    

def  excel_process():
    args=Args()
    book=Book(args.filepath,args.sheets)
    book.apply_rules(args.exportpath,args.rules)
    book.close()
    
    return 
if __name__=="__main__":
    
    excel_process()

