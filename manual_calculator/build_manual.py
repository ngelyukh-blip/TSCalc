"""Build the manual TSCalc workbook; legacy source files are never modified."""
from pathlib import Path
import sys, hashlib, json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import FormulaRule
from openpyxl.workbook.properties import CalcProperties
from openpyxl.workbook.defined_name import DefinedName

ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/'TSCalc_Manual_v0.2.xlsx'
LEND=2001; PEND=1001; OEND=1001; MEND=301; NEND=101
wb=openpyxl.Workbook(); wb.remove(wb.active)
names=['Инструкция','Итог','Лаги','Щиты','Прочие','Материалы','Узлы','Расчёт','Спецификация','Проверки']
for n in names: wb.create_sheet(n)
wb.calculation=CalcProperties(calcId=191029,fullCalcOnLoad=True,forceFullCalc=True)
navy='17365D'; teal='DDEBF7'; yellow='FFF2CC'; gray='E7E6E6'; blue='0000CC'

def init(name,headers,widths,inputs=()):
 s=wb[name] if name in wb.sheetnames else wb.create_sheet(name);s.append(headers);s.freeze_panes='A2';s.sheet_view.showGridLines=False
 for c in s[1]: c.fill=PatternFill('solid',fgColor=navy);c.font=Font(name='Calibri',size=11,bold=True,color='FFFFFF');c.alignment=Alignment(wrap_text=True,vertical='center')
 s.row_dimensions[1].height=42
 from openpyxl.utils import get_column_letter
 for i,w in enumerate(widths,1):s.column_dimensions[get_column_letter(i)].width=w
 s.sheet_properties.pageSetUpPr.fitToPage=True;s.page_setup.orientation='landscape';s.page_setup.paperSize=s.PAPERSIZE_A3;s.page_setup.fitToWidth=1;s.page_setup.fitToHeight=0;s.print_title_rows='1:1'
 s.sheet_properties.tabColor='4472C4' if inputs else '548235'
 return s

def finish(s,end,inputcols,table=None):
 lastcol=s.max_column
 for row in s.iter_rows(min_row=2,max_row=end):
  for c in row:
   c.font=Font(name='Calibri',size=10,color=blue if c.column in inputcols else '222222')
   c.fill=PatternFill('solid',fgColor=yellow if c.column in inputcols else 'F2F2F2')
   c.alignment=Alignment(vertical='center',wrap_text=c.column==lastcol)
 s.auto_filter.ref=f'A1:{openpyxl.utils.get_column_letter(s.max_column)}{end}'
 if table:
  t=Table(displayName=table,ref=s.auto_filter.ref);t.tableStyleInfo=TableStyleInfo(name='TableStyleMedium2',showRowStripes=True);s.add_table(t)
 s.print_options.horizontalCentered=True;s.sheet_properties.outlinePr.summaryRight=False

def dv(s,typ,formula,rng,op=None,f2=None):
 d=DataValidation(type=typ,formula1=formula,formula2=f2,operator=op,allow_blank=True);d.showErrorMessage=True;d.errorTitle='Проверьте ввод';d.error='Выберите допустимое значение. Числа должны соответствовать ограничению.';d.errorStyle='stop';s.add_data_validation(d);d.add(rng)

s=init('Материалы',['ID','Группа','Наименование / порода','Ед. расчёта','Длина изделия, мм','Ширина, мм','Высота / толщина, мм','Кратность в ед. расчёта','Коэффициент запаса','Цена за ед. расчёта, руб.','Статус характеристик','Источник / дата цены'],[19,18,54,16,17,15,20,22,21,23,23,45],range(1,13))
catalog=[]
for i,(l,b,h) in enumerate([(2400,1200,60),(2400,1200,70),(2000,1000,60)],1):
 for j,wood in enumerate(['Ангарская сосна','Лиственница'],1):catalog.append([f'Щ-{i}{j}','Щиты',f'Щит {l}×{b}×{h}, {wood}','шт',l,b,h,1,1,None,'Подтверждено','Размеры и породы: указание пользователя; цену заполнить'])
for i,(b,h) in enumerate([(100,100),(100,150),(85,85),(95,95)],1):catalog.append([f'ЛАГ-{i:03}','Лаги',f'Брус лаг {b}×{h}; укажите породу','м³',None,b,h,0.001,1,None,'Не подтверждено','Пример сечения из старого шаблона; подтвердить/изменить'])
for id,group,name,unit,pack in [('КР-001','Крепёж','Крепёж 1 — задайте изделие','шт',1),('КР-002','Крепёж','Крепёж 2 — задайте изделие','шт',1),('КР-003','Крепёж','Крепёж 3 — задайте изделие','шт',1),('ВОЙ-001','Прокладки','Войлок — задайте характеристики','м²',1),('ПЛ-001','Плинтус','Плинтус — задайте характеристики','м',1),('РЕЙ-001','Рейка','Рейка дубовая — задайте сечение','м',1),('КЛ-001','Клей','Клей — задайте продукт','кг',1),('ОЗ-001','Покрытия','Огнезащита — задайте продукт','кг',1)]:catalog.append([id,group,name,unit,None,None,None,pack,1,None,'Не подтверждено','Заполнить характеристики и источник'])
for row in catalog:s.append(row)
finish(s,MEND,range(1,13),'Materials')
dv(s,'list','"Подтверждено,Не подтверждено"',f'K2:K{MEND}');dv(s,'list','"шт,м,м²,м³,кг,л"',f'D2:D{MEND}')
dv(s,'decimal','0',f'H2:H{MEND}','greaterThan');dv(s,'decimal','1',f'I2:I{MEND}','greaterThanOrEqual');dv(s,'decimal','0',f'J2:J{MEND}','greaterThanOrEqual')
wb.defined_names.add(DefinedName('MaterialIDs',attr_text=f"'Материалы'!$A$2:$A${MEND}"))
wb.defined_names.add(DefinedName('NodeIDs',attr_text=f"'Узлы'!$A$2:$A${NEND}"))

s=init('Узлы',['ID узла','Основание','Метод точек на кусок','Шаг, мм','Статус узла','Источник / правило концов','Материал 1 ID','На точку 1','Материал 2 ID','На точку 2','Материал 3 ID','На точку 3'],[18,18,24,16,23,58,20,17,20,17,20,17],range(1,13))
for row in [['БЕТ-01','Бетон','Вручную',None,'Не подтверждено','Введите подтверждённые точки; комплект задайте справа'],['МЕТ-01','Металл','Вручную',None,'Не подтверждено','Введите подтверждённые точки; комплект задайте справа'],['ДЕР-01','Дерево','Вручную',None,'Не подтверждено','Пересечения ярусов задаются вручную'],['ВИБ-01','Бетон','Вручную',None,'Не подтверждено','Виброузел: основание и комплект уточнить по проекту']]:s.append(row)
finish(s,NEND,range(1,13),'Nodes')
dv(s,'list','"Бетон,Металл,Дерево"',f'B2:B{NEND}');dv(s,'list','"Вручную,По шагу"',f'C2:C{NEND}');dv(s,'list','"Подтверждено,Не подтверждено"',f'E2:E{NEND}');dv(s,'decimal','0',f'D2:D{NEND}','greaterThan')
for c in ['G','I','K']:dv(s,'list','MaterialIDs',f'{c}2:{c}{NEND}')
for c in ['H','J','L']:dv(s,'decimal','0',f'{c}2:{c}{NEND}','greaterThan')

def lookup(key,col,table='Материалы',end=MEND):return f'IFERROR(VLOOKUP({key},\'{table}\'!$A$2:$L${end},{col},FALSE),"")'
def mget(key,col):return lookup(key,col)
def nget(key,col):return lookup(key,col,'Узлы',NEND)

s=init('Лаги',['ID строки','Материал лаги ID','Длина одного куска, мм','Кусков, шт','Основание','Узел ID','Точек на кусок вручную','Подтверждение количества','Источник / зона','Ширина, мм','Высота, мм','Длина всего, м','Объём чисто, м³','Точек всего','Материал 1 ID','Кол-во 1','Материал 2 ID','Кол-во 2','Материал 3 ID','Кол-во 3','Статус'],[20,20,25,16,18,18,26,27,38,16,16,19,21,18,20,18,20,18,20,18,52],range(1,10))
for r in range(2,LEND+1):
 k=f'B{r}';n=f'F{r}';present=f'COUNTA(A{r}:I{r})=0';method=nget(n,3);step=nget(n,4)
 status=f'IF({present},"",IF(OR(A{r}="",COUNTIF($A$2:$A${LEND},A{r})<>1),"Ошибка: ID строки",IF(OR(COUNTIF(MaterialIDs,B{r})<>1,{mget(k,2)}<>"Лаги",{mget(k,4)}<>"м³"),"Ошибка: материал лаги",IF(OR(NOT(ISNUMBER(C{r})),C{r}<=0,NOT(ISNUMBER(D{r})),D{r}<=0,MOD(D{r},1)<>0),"Ошибка: длина / куски",IF(OR(J{r}<=0,K{r}<=0),"Ошибка: сечение",IF(OR(COUNTIF(NodeIDs,F{r})<>1,E{r}<>{nget(n,2)}),"Ошибка: основание / узел",IF(OR(H{r}<>"Подтверждено",{nget(n,5)}<>"Подтверждено",{mget(k,11)}<>"Подтверждено"),"Не подтверждены исходные данные",IF(AND({method}="Вручную",OR(NOT(ISNUMBER(G{r})),G{r}<0,MOD(G{r},1)<>0)),"Ошибка: ручные точки",IF(AND({method}="По шагу",OR(NOT(ISNUMBER({step})),{step}<=0)),"Ошибка: шаг",IF(AND({method}<>"Вручную",{method}<>"По шагу"),"Ошибка: метод","ОК")))))))))'
 s.cell(r,10,f'={mget(k,6)}');s.cell(r,11,f'={mget(k,7)}');s.cell(r,21,'='+status)
 s.cell(r,12,f'=IF(U{r}="ОК",C{r}*D{r}/1000,0)');s.cell(r,13,f'=IF(U{r}="ОК",C{r}*D{r}*J{r}*K{r}/1000000000,0)')
 s.cell(r,14,f'=IF(U{r}="ОК",D{r}*IF({method}="Вручную",G{r},ROUNDUP(C{r}/IF({step}>0,{step},1),0)+1),0)')
 for idcol,qtycol,nid,nqty in [(15,16,7,8),(17,18,9,10),(19,20,11,12)]:
  s.cell(r,idcol,f'=IF(U{r}="ОК",{nget(n,nid)},"")');s.cell(r,qtycol,f'=IF(U{r}="ОК",N{r}*N({nget(n,nqty)}),0)')
finish(s,LEND,range(1,10),'LagInput')
dv(s,'list','MaterialIDs',f'B2:B{LEND}');dv(s,'list','NodeIDs',f'F2:F{LEND}');dv(s,'list','"Бетон,Металл,Дерево"',f'E2:E{LEND}');dv(s,'list','"Подтверждено,Не подтверждено"',f'H2:H{LEND}')
dv(s,'decimal','0',f'C2:C{LEND}','greaterThan');dv(s,'whole','1',f'D2:D{LEND}','greaterThanOrEqual');dv(s,'whole','0',f'G2:G{LEND}','greaterThanOrEqual')

s=init('Щиты',['ID строки','Материал щита ID','Количество, шт','Подтверждение количества','Источник / зона','Длина, мм','Ширина, мм','Толщина, мм','Площадь чисто, м²','Объём, м³','Статус'],[20,22,20,28,42,18,18,18,24,20,52],range(1,6))
for r in range(2,PEND+1):
 for c,mc in [(6,5),(7,6),(8,7)]:s.cell(r,c,f'={mget(f"B{r}",mc)}')
 s.cell(r,11,f'=IF(COUNTA(A{r}:E{r})=0,"",IF(OR(A{r}="",COUNTIF($A$2:$A${PEND},A{r})<>1),"Ошибка: ID строки",IF(OR(COUNTIF(MaterialIDs,B{r})<>1,{mget(f"B{r}",2)}<>"Щиты",{mget(f"B{r}",4)}<>"шт",AND(B{r}<>"Щ-11",B{r}<>"Щ-12",B{r}<>"Щ-21",B{r}<>"Щ-22",B{r}<>"Щ-31",B{r}<>"Щ-32")),"Ошибка: каталог щитов",IF(OR(NOT(ISNUMBER(C{r})),C{r}<=0,MOD(C{r},1)<>0),"Ошибка: количество",IF(OR(D{r}<>"Подтверждено",{mget(f"B{r}",11)}<>"Подтверждено"),"Не подтверждены исходные данные","ОК")))))')
 s.cell(r,9,f'=IF(K{r}="ОК",C{r}*F{r}*G{r}/1000000,0)');s.cell(r,10,f'=IF(K{r}="ОК",I{r}*H{r}/1000,0)')
finish(s,PEND,range(1,6),'PanelInput');dv(s,'list','MaterialIDs',f'B2:B{PEND}');dv(s,'whole','1',f'C2:C{PEND}','greaterThanOrEqual');dv(s,'list','"Подтверждено,Не подтверждено"',f'D2:D{PEND}')

s=init('Прочие',['ID строки','Материал ID','Чистая потребность, ед. расчёта','Подтверждение количества','Источник / расчёт / узел','Ед. расчёта','Статус'],[20,22,31,28,64,19,52],range(1,6))
for r in range(2,OEND+1):
 s.cell(r,6,f'={mget(f"B{r}",4)}')
 s.cell(r,7,f'=IF(COUNTA(A{r}:E{r})=0,"",IF(OR(A{r}="",COUNTIF($A$2:$A${OEND},A{r})<>1),"Ошибка: ID строки",IF(COUNTIF(MaterialIDs,B{r})<>1,"Ошибка: материал",IF(OR(NOT(ISNUMBER(C{r})),C{r}<0,AND(F{r}="шт",MOD(C{r},1)<>0)),"Ошибка: количество",IF(OR(D{r}<>"Подтверждено",E{r}="",{mget(f"B{r}",11)}<>"Подтверждено"),"Не подтверждены исходные данные","ОК")))))')
finish(s,OEND,range(1,6),'OtherInput');dv(s,'list','MaterialIDs',f'B2:B{OEND}');dv(s,'decimal','0',f'C2:C{OEND}','greaterThanOrEqual');dv(s,'list','"Подтверждено,Не подтверждено"',f'D2:D{OEND}')

s=init('Расчёт',['ID материала','Наименование','Ед.','Чистая потребность','Коэф. запаса','Запас, ед.','С запасом','Кратность','Закупка, ед.','Упаковок / кратностей','Цена за ед., руб.','Стоимость, руб.','Статус'],[20,54,14,23,18,20,21,19,23,25,24,25,49])
for r in range(2,MEND+1):
 s.cell(r,1,f'=IF(Материалы!A{r}="","",Материалы!A{r})');s.cell(r,2,f'=IF(A{r}="","",Материалы!C{r})');s.cell(r,3,f'=IF(A{r}="","",Материалы!D{r})')
 s.cell(r,4,f'=IF(A{r}="",0,SUMIF(Лаги!$B$2:$B${LEND},A{r},Лаги!$M$2:$M${LEND})+SUMIFS(Щиты!$C$2:$C${PEND},Щиты!$B$2:$B${PEND},A{r},Щиты!$K$2:$K${PEND},"ОК")+SUMIFS(Прочие!$C$2:$C${OEND},Прочие!$B$2:$B${OEND},A{r},Прочие!$G$2:$G${OEND},"ОК")+SUMIF(Лаги!$O$2:$O${LEND},A{r},Лаги!$P$2:$P${LEND})+SUMIF(Лаги!$Q$2:$Q${LEND},A{r},Лаги!$R$2:$R${LEND})+SUMIF(Лаги!$S$2:$S${LEND},A{r},Лаги!$T$2:$T${LEND}))')
 s.cell(r,5,f'=IF(A{r}="","",Материалы!I{r})');s.cell(r,8,f'=IF(A{r}="","",Материалы!H{r})');s.cell(r,11,f'=IF(Материалы!J{r}="","",Материалы!J{r})')
 s.cell(r,13,f'=IF(A{r}="","",IF(COUNTIF(MaterialIDs,A{r})<>1,"Ошибка: дубликат ID",IF(D{r}=0,"Нет потребности",IF(OR(NOT(ISNUMBER(E{r})),E{r}<1,NOT(ISNUMBER(H{r})),H{r}<=0),"Ошибка: запас / кратность",IF(Материалы!K{r}<>"Подтверждено","Не подтверждён материал",IF(K{r}="","Нет цены",IF(OR(NOT(ISNUMBER(K{r})),K{r}<0),"Ошибка: цена","ОК")))))))')
 s.cell(r,7,f'=IF(AND(D{r}>0,ISNUMBER(E{r}),E{r}>=1),D{r}*E{r},0)');s.cell(r,6,f'=IF(G{r}>0,G{r}-D{r},0)')
 s.cell(r,9,f'=IF(AND(G{r}>0,ISNUMBER(H{r}),H{r}>0),ROUNDUP(ROUND(G{r}/IF(H{r}>0,H{r},1),10),0)*H{r},0)');s.cell(r,10,f'=IF(AND(I{r}>0,H{r}>0),I{r}/IF(H{r}>0,H{r},1),0)');s.cell(r,12,f'=IF(M{r}="ОК",ROUND(I{r}*K{r},2),0)')
finish(s,MEND,[],'Calculation')
for row in s.iter_rows(min_row=2):
 for c in row:
  if 4<=c.column<=12:c.number_format='#,##0.000' if c.column not in [11,12] else '#,##0.00'

s=init('Спецификация',['ID материала','Наименование','Ед.','Чисто','Запас','С запасом','Закупка','Цена, руб.','Стоимость, руб.','Статус'],[20,54,14,22,20,22,22,22,24,50])
for r in range(2,MEND+1):
 for c,src in [(1,'A'),(2,'B'),(3,'C'),(4,'D'),(5,'F'),(6,'G'),(7,'I'),(8,'K'),(9,'L'),(10,'M')]:s.cell(r,c,f'=IF(Расчёт!D{r}>0,Расчёт!{src}{r},"")')
finish(s,MEND,[],'Specification')
s['L1']='Отфильтруйте пустые ID';s.column_dimensions['L'].width=30

s=init('Проверки',['Проверка','Количество / результат','Действие'],[52,30,90])
checks=[('Ошибки / неподтверждённые лаги',f'=COUNTIF(Лаги!U2:U{LEND},"Ошибка*")+COUNTIF(Лаги!U2:U{LEND},"Не подтверждены*")','Фильтр «Статус» на листе «Лаги».'),('Ошибки / неподтверждённые щиты',f'=COUNTIF(Щиты!K2:K{PEND},"Ошибка*")+COUNTIF(Щиты!K2:K{PEND},"Не подтверждены*")','Фильтр «Статус» на листе «Щиты».'),('Ошибки / неподтверждённые прочие',f'=COUNTIF(Прочие!G2:G{OEND},"Ошибка*")+COUNTIF(Прочие!G2:G{OEND},"Не подтверждены*")','Прочие потребности — чистые, без запаса. Не дублировать автоматический крепёж.'),('Позиции без цены',f'=COUNTIF(Расчёт!M2:M{MEND},"Нет цены")','Количество посчитано; сумма неполная до заполнения цен.'),('Ошибки материалов в потребности',f'=COUNTIF(Расчёт!M2:M{MEND},"Ошибка*")+COUNTIF(Расчёт!M2:M{MEND},"Не подтверждён*")','Проверьте ID, характеристики, запас, кратность и цену.'),('Строки лаг за пределом формы',f'=COUNTA(Лаги!A:A)-COUNTA(Лаги!A1:A{LEND})','Максимум 2000 строк. За пределом не рассчитывается; требуется расширение формы.'),('Строки щитов за пределом формы',f'=COUNTA(Щиты!A:A)-COUNTA(Щиты!A1:A{PEND})','Максимум 1000 строк. За пределом не рассчитывается.'),('Строки прочих за пределом формы',f'=COUNTA(Прочие!A:A)-COUNTA(Прочие!A1:A{OEND})','Максимум 1000 строк. За пределом не рассчитывается.'),('Материалы за пределом каталога',f'=COUNTA(Материалы!A:A)-COUNTA(Материалы!A1:A{MEND})','Максимум 300 материалов.'),('Узлы за пределом каталога',f'=COUNTA(Узлы!A:A)-COUNTA(Узлы!A1:A{NEND})','Максимум 100 узлов.'),('Введено строк',f'=COUNTA(Лаги!A2:A{LEND})+COUNTA(Щиты!A2:A{PEND})+COUNTA(Прочие!A2:A{OEND})','Пустая форма не является завершённым расчётом.')]
for row in checks:s.append(row)
finish(s,len(checks)+1,[])

s=init('Итог',['Показатель','Значение','Пояснение'],[49,33,95])
for row in [('Объект / заказ',None,'Введите наименование объекта.'),('Версия','TSCalc Manual 0.2','Ручной ввод; логика DWG не включена.'),('Готовность количеств','=IF(Проверки!B12=0,"ПУСТАЯ ФОРМА",IF(SUM(Проверки!B2:B4,Проверки!B6:B11)>0,"ЕСТЬ ЗАМЕЧАНИЯ","КОЛИЧЕСТВА ПРОВЕРЕНЫ"))','Статус проверяет заполнение, а не конструктивную безопасность проекта.'),('Готовность стоимости','=IF(B4<>"КОЛИЧЕСТВА ПРОВЕРЕНЫ","НЕ ГОТОВО",IF(Проверки!B5>0,"НЕПОЛНАЯ СТОИМОСТЬ","ЦЕНЫ ЗАПОЛНЕНЫ"))','Стоимость только материалов; цена за базовую единицу, НДС/доставка/работы отдельно не добавляются.'),('Сумма известных стоимостей, руб.',f'=SUM(Расчёт!L2:L{MEND})','При пропущенных ценах это частичный итог, не окончательная стоимость.'),('Длина лаг чисто, м',f'=SUM(Лаги!L2:L{LEND})','Только строки со статусом ОК.'),('Объём лаг чисто, м³',f'=SUM(Лаги!M2:M{LEND})','Только строки со статусом ОК.'),('Площадь полных щитов, м²',f'=SUM(Щиты!I2:I{PEND})','Форма считает полные изделия; обрезки и раскрой не моделируются.'),('Дата выпуска',None,'Заполните дату после проверки исходных данных.')]:s.append(row)
finish(s,11,[2]);s['B2'].fill=PatternFill('solid',fgColor=yellow);s['B11'].fill=PatternFill('solid',fgColor=yellow);s['B6'].number_format='#,##0.00'

s=wb['Инструкция'];s.sheet_view.showGridLines=False;s.column_dimensions['A'].width=8;s.column_dimensions['B'].width=115
instructions=[('TSCalc Manual 0.2','Калькулятор материалов по ручной ведомости'),('1. Объект','Заполните объект на листе «Итог». Исходный старый Excel не изменён.'),('2. Материалы','Шесть сочетаний щитов готовы. У лаг и прочих материалов заполните характеристики и подтвердите их. Цена — за единицу расчёта. Для цены за упаковку предварительно разделите её на содержимое упаковки.'),('3. Запас и упаковка','Кратность — количество в единицах расчёта; коэффициент 1 означает без запаса. Запас применяется один раз после объединения потребности одного ID по объекту. Цены и нормы не перенесены как утверждённые.'),('4. Узлы','Задайте основание, метод подсчёта точек и состав комплекта. «Вручную»: число точек на один физический кусок. «По шагу»: ceil(длина куска/шаг)+1, с точками у обоих концов каждого куска; применять только если это правило подтверждено проектом. Нулевые точки разрешены при явном ручном вводе.'),('5. Лаги','Одна строка: одинаковые физические куски одной длины и сечения. Введите число кусков, основание, узел, источник и подтверждение. Не вводите общую линию как кусок при нескольких реальных отрезках. Стыки, пластины и крепления ярусов при отсутствии утверждённого узла задайте через «Прочие».'),('6. Щиты','Выберите ID полного щита и целое количество. Типоразмеры только 2400×1200×60, 2400×1200×70, 2000×1000×60; материал ангарская сосна / лиственница. Породы и толщины закупаются раздельно. Обрезки/раскрой сегодня вне модели — полные изделия вводятся по проверенной ведомости закупки.'),('7. Прочие','Вводите чистую потребность войлока, клея, огнезащиты, плинтуса, рейки, шпонок, заглушек и прочих материалов с источником нормы/ручного расчёта. Автоматических универсальных норм для них нет. Не дублируйте крепёж, уже включённый в узлы лаг.'),('8. Подтверждение','«Подтверждено» означает, что пользователь проверил исходное количество/характеристики/узел по проекту или паспорту. Это не автоматическое нормативное заключение. Формулы и книга не заменяют проверку конструкции.'),('9. Проверка','Все ошибки/неподтверждённые строки видны в «Проверки» и на листах ввода. Результаты по корректным строкам частичные при любых замечаниях. Пустая цена не равна нулю. Нулевая цена допустима только осознанно.'),('10. Вывод','«Расчёт» содержит каждую позицию каталога, «Спецификация» — строки с потребностью. Отфильтруйте пустые ID перед печатью. Количество упаковок и закупка показаны отдельно. Стоимость округляется до копеек по позиции.'),('Ограничения формы','2000 строк лаг,1000 строк щитов,1000 строк прочих,300 материалов,100 узлов. Не вводите ниже границы: индикатор проверяет ID за пределом. Вставка/расширение требует проверки формул; продолжение работы — через разработчика. Не удаляйте/переименовывайте служебные листы.'),('Цвета','Жёлтая заливка и синий шрифт — ввод; серый фон — формулы. Таблицы можно фильтровать. Формулы пересчитываются при открытии Excel; проверяйте отсутствие замечаний перед выпуском.'),('Цена / состав','Книга считает стоимость материалов в рублях по введённым ценам. Укажите в источнике дату и включение НДС. Монтаж, доставка, накладные, прибыль и отдельное начисление НДС не реализованы без согласованной структуры КП.')]
for i,(a,b) in enumerate(instructions,1):s.cell(i,1,i);s.cell(i,2,a+'\n'+b);s.cell(i,2).alignment=Alignment(wrap_text=True,vertical='top');s.cell(i,2).font=Font(name='Calibri',size=12,bold=i==1,color=navy);s.row_dimensions[i].height=70 if i in [4,5,6,7,8,13] else 55
s.freeze_panes='B2';s.print_area=f'B1:B{len(instructions)}';s.page_setup.orientation='portrait';s.page_setup.paperSize=s.PAPERSIZE_A4;s.page_setup.fitToWidth=1;s.page_setup.fitToHeight=0;s.sheet_properties.pageSetUpPr.fitToPage=True
## Automatic legacy material consumption, driven by manual geometry and zones.
legacy=openpyxl.load_workbook((Path(sys.argv[1]) if len(sys.argv)>1 else ROOT/'template.xlsx'),data_only=False)
ps=wb.create_sheet('Параметры');oldps=legacy['Параметры']
from copy import copy
for row in oldps:
 for c in row:
  ps.cell(c.row,c.column,c.value);ps.cell(c.row,c.column)._style=copy(c._style)
ps.column_dimensions['A'].width=64;ps.column_dimensions['B'].width=26;ps.column_dimensions['C'].width=16;ps.column_dimensions['D'].width=78;ps.column_dimensions['E'].width=22
ps.freeze_panes='B3';ps['B3']=None
ps['A100']='Нормы подтверждены для этого объекта';ps['B100']='Нет';ps['D100']='Исторические параметры 0.1 сохранены как НАСТРОЙКИ. До проверки проекта/паспортов расчёт черновой.'
ps['A102']='Саморезов на полный щит';ps['B102']=15;ps['C102']='шт';ps['D102']='Настройка старого примера; проверить по выбранному щиту и узлу.'
dv(ps,'list','"Да,Нет"','B100');dv(ps,'whole','0','B102','greaterThanOrEqual')
for key,dn in legacy.defined_names.items():wb.defined_names.add(DefinedName(key,attr_text=dn.attr_text))
zs=init('Зоны',['ID зоны','Основание','Площадь, м²','Виброизоляция','Периметр акустического шва, м','Источник / примечание','Статус'],[22,18,22,22,35,60,45],range(1,7))
for r in range(2,202):
 zs.cell(r,7,f'=IF(COUNTA(A{r}:F{r})=0,"",IF(OR(A{r}="",COUNTIF($A$2:$A$201,A{r})<>1,AND(B{r}<>"Бетон",B{r}<>"Металл",B{r}<>"Дерево"),NOT(ISNUMBER(C{r})),C{r}<=0,AND(D{r}<>"Да",D{r}<>"Нет"),NOT(ISNUMBER(E{r})),E{r}<0),"Ошибка: зона / площадь / периметр","ОК"))')
finish(zs,201,range(1,7),'ZoneInput');dv(zs,'list','"Бетон,Металл,Дерево"','B2:B201');dv(zs,'list','"Да,Нет"','D2:D201')
ls=init('Рейка и плинтус',['ID строки','Вид','Длина одного отрезка, мм','Количество отрезков','Зона ID','Длина всего, м','Статус'],[22,23,31,25,22,25,50],range(1,6))
for r in range(2,1002):
 ls.cell(r,7,f'=IF(COUNTA(A{r}:E{r})=0,"",IF(OR(A{r}="",COUNTIF($A$2:$A$1001,A{r})<>1,AND(B{r}<>"Плинтус",B{r}<>"Дубовая рейка"),NOT(ISNUMBER(C{r})),C{r}<=0,NOT(ISNUMBER(D{r})),D{r}<=0,MOD(D{r},1)<>0,COUNTIF(Зоны!$A$2:$A$201,E{r})<>1),"Ошибка: длина / количество / зона","ОК"))')
 ls.cell(r,6,f'=IF(G{r}="ОК",C{r}*D{r}/1000,0)')
finish(ls,1001,range(1,6),'LinearInput');dv(ls,'list','"Плинтус,Дубовая рейка"','B2:B1001')
ns=wb['Узлы']
for r,stepname in [(2,'STEP_CONC'),(3,'STEP_MET'),(5,'STEP_VIB')]:ns.cell(r,3,'По шагу');ns.cell(r,4,'='+stepname)
ns['B5']='Бетон';ns['F5']='Виброузел отдельно: узел ВИБ-01. Шаг из настроек; проверить паспорт.'
# Lag inputs: physical pieces; quantities are explicit user input, no per-row confirmation.
lags=wb['Лаги'];lags['H1']='Зона ID';lags['V1']='Стыков физически, шт';lags.column_dimensions['V'].width=25
for r in range(2,LEND+1):
 f=lags.cell(r,21).value.replace(f'H{r}<>"Подтверждено",','').replace(f'{nget(f"F{r}",5)}<>"Подтверждено",','').replace(f'{mget(f"B{r}",11)}<>"Подтверждено"','FALSE')
 f=f.replace(f'J{r}<=0,K{r}<=0',f'NOT(ISNUMBER(J{r})),NOT(ISNUMBER(K{r})),J{r}<=0,K{r}<=0')
 f=f.replace('"ОК"))))',f'IF(OR(COUNTIF(Зоны!$A$2:$A$201,H{r})<>1,NOT(ISNUMBER(V{r})),V{r}<0,MOD(V{r},1)<>0,AND(W{r}<>"Да",W{r}<>"Нет")),"Ошибка: зона / стыки / войлок","ОК")))))')
 lags.cell(r,21,f+')');lags.cell(r,22).fill=PatternFill('solid',fgColor=yellow);lags.cell(r,22).font=Font(name='Calibri',color=blue,size=10)
# Remove old row-confirmation validation (H is now a zone ID).
lags.data_validations.dataValidation=[v for v in lags.data_validations.dataValidation if not str(v.sqref).startswith('H2:')]
panels=wb['Щиты'];panels['D1']='Зона ID'
for r in range(2,PEND+1):
 f=panels.cell(r,11).value.replace(f'D{r}<>"Подтверждено"',f'COUNTIF(Зоны!$A$2:$A$201,D{r})<>1').replace(f'{mget(f"B{r}",11)}<>"Подтверждено"','FALSE')
 panels.cell(r,11,f)
panels.data_validations.dataValidation=[v for v in panels.data_validations.dataValidation if not str(v.sqref).startswith('D2:')]
# All old named totals are rebound to the expanded, explicit geometry input.
rebind={
'A_TOTAL':'SUMIFS(Зоны!C2:C201,Зоны!G2:G201,"ОК")',
'A_CONC':'SUMIFS(Зоны!C2:C201,Зоны!B2:B201,"Бетон",Зоны!G2:G201,"ОК")',
'A_VIB':'SUMIFS(Зоны!C2:C201,Зоны!D2:D201,"Да",Зоны!G2:G201,"ОК")',
'PER_ROOM':'SUMIFS(Зоны!E2:E201,Зоны!D2:D201,"Да",Зоны!G2:G201,"ОК")',
'PER_MOV':'SUMIFS(\'Рейка и плинтус\'!F2:F1001,\'Рейка и плинтус\'!B2:B1001,"Дубовая рейка")',
'L_CONC':f'SUMIFS(Лаги!L2:L{LEND},Лаги!E2:E{LEND},"Бетон",Лаги!F2:F{LEND},"<>ВИБ-01")',
'L_MET':f'SUMIFS(Лаги!L2:L{LEND},Лаги!E2:E{LEND},"Металл",Лаги!F2:F{LEND},"<>ВИБ-01")',
'L_LAG':f'SUMIFS(Лаги!L2:L{LEND},Лаги!E2:E{LEND},"Дерево",Лаги!F2:F{LEND},"<>ВИБ-01")',
'L_VIB':f'SUMIF(Лаги!F2:F{LEND},"ВИБ-01",Лаги!L2:L{LEND})',
'P_CONC':f'SUMIFS(Лаги!N2:N{LEND},Лаги!E2:E{LEND},"Бетон",Лаги!F2:F{LEND},"<>ВИБ-01")',
'P_MET':f'SUMIFS(Лаги!N2:N{LEND},Лаги!E2:E{LEND},"Металл",Лаги!F2:F{LEND},"<>ВИБ-01")',
'P_LAG':f'SUMIFS(Лаги!N2:N{LEND},Лаги!E2:E{LEND},"Дерево",Лаги!F2:F{LEND},"<>ВИБ-01")',
'P_VIB':f'SUMIF(Лаги!F2:F{LEND},"ВИБ-01",Лаги!N2:N{LEND})',
'JOINTS':f'SUMIFS(Лаги!V2:V{LEND},Лаги!U2:U{LEND},"ОК")',
'N_PANELS':f'SUMIFS(Щиты!C2:C{PEND},Щиты!K2:K{PEND},"ОК")',
'A_PANELS':f'SUM(Щиты!I2:I{PEND})',
'PER_PANELS':f'SUMPRODUCT(Щиты!C2:C{PEND},Щиты!F2:F{PEND}+Щиты!G2:G{PEND},--(Щиты!K2:K{PEND}="ОК"))*2/1000',
'EDGE_A':f'SUMPRODUCT(Щиты!C2:C{PEND},Щиты!F2:F{PEND}+Щиты!G2:G{PEND},Щиты!H2:H{PEND},--(Щиты!K2:K{PEND}="ОК"))*2/1000000',
'SCR6':'N_PANELS*Параметры!B102',
'FELT_A':f'SUMPRODUCT(Лаги!L2:L{LEND},Лаги!J2:J{LEND})/1000',
'PAINT_LAG_A':f'SUMPRODUCT(Лаги!L2:L{LEND},Лаги!J2:J{LEND}+Лаги!K2:K{LEND})*2/1000',
'MET_STRIP':f'SUMPRODUCT(Лаги!L2:L{LEND},Лаги!J2:J{LEND},--(Лаги!E2:E{LEND}="Металл"),--(Лаги!F2:F{LEND}<>"ВИБ-01"))/1000',
'VIB_STRIP':f'SUMPRODUCT(Лаги!L2:L{LEND},Лаги!J2:J{LEND},--(Лаги!F2:F{LEND}="ВИБ-01"))/1000',
'HMAX_CONC':f'MAX(Лаги!K2:K{LEND})','HMAX_MET':f'MAX(Лаги!K2:K{LEND})','T_MAX':f'MAX(Щиты!H2:H{PEND})'}
for i in range(1,5):
 rebind[f'V_S{i}']=f'SUMIF(Лаги!B2:B{LEND},"ЛАГ-{i:03}",Лаги!M2:M{LEND})';rebind[f'LEN_S{i}']=f'SUMIF(Лаги!B2:B{LEND},"ЛАГ-{i:03}",Лаги!L2:L{LEND})'
for key,formula in rebind.items():
 addr=legacy.defined_names[key].attr_text.split('$B$')[-1];ps[f'B{addr}']='='+formula
ps['A7']='Периметр акустического шва виброзон';ps['A8']='Длина дубовой рейки';ps['B9']=0;ps['D9']='Наружный контур щитового поля задайте вручную; не сумма рейки и акустического периметра.'
# Bind lag cross-sections to user settings, avoiding duplicated settings.
for i in range(1,5):
 rr=7+i;wb['Материалы'].cell(rr,6,f'=S{i}_B');wb['Материалы'].cell(rr,7,f'=S{i}_H')
# Recover the full legacy list, replacing its three panel positions with all six catalog IDs.
auto=[]
for r in range(2,33):
 if r in [21,22,23]:continue
 auto.append([legacy['Расчёт'].cell(r,c).value for c in range(1,12)])
ms=wb['Материалы'];cs=wb['Расчёт']
for idx,data in enumerate(auto,30):
 code,group,label,unit,formula,reserve,_,pack,_,explain,_=data
 ms.cell(idx,1,code);ms.cell(idx,2,group);ms.cell(idx,3,label);ms.cell(idx,4,unit);ms.cell(idx,8,pack);ms.cell(idx,9,reserve);ms.cell(idx,11,'Подтверждено');ms.cell(idx,12,'Настройка / номенклатура старого Excel; нормы требуют подтверждения на листе Параметры.')
 if code=='ПЛИНТ':formula='=SUMIFS(\'Рейка и плинтус\'!F2:F1001,\'Рейка и плинтус\'!B2:B1001,"Плинтус")'
 if code=='ПЛС':formula='=IF(PLINTH_STEP>0,SUMIFS(\'Рейка и плинтус\'!F2:F1001,\'Рейка и плинтус\'!B2:B1001,"Плинтус")/PLINTH_STEP,0)'
 if code=='ХИМ':formula='=IF(STUDS_CART>0,P_CONC/STUDS_CART,0)'
 if code=='ЛЕНТА':formula='=IF(AND(L_VIB>0,TAPE_ROLL>0),(PER_ROOM+IF(TAPE_UNDER="Да",L_VIB,0))/TAPE_ROLL,0)'
 if code=='ГЕРМ':formula='=IF(AND(L_VIB>0,SEAL_M>0),PER_ROOM/SEAL_M,0)'
 if code=='ШПОН':formula='=IF(SHPON_L>0,MAX(0,PER_PANELS-CONTOUR)/2/SHPON_L,0)'
 cs.cell(idx,4,formula+f'+SUMIFS(Прочие!$C$2:$C${OEND},Прочие!$B$2:$B${OEND},A{idx},Прочие!$G$2:$G${OEND},"ОК")')
# Geometry-only first ten catalog rows retain their aggregation. Auto timber is represented by БРУС1..4 only.
for r in range(8,12):cs.cell(r,4,0)
for r in range(2,8):ms.cell(r,9,'=R_PANEL')
# No per-row approvals for normal manual quantities. Optional additional needs remain explicitly marked.
os=wb['Прочие'];os['D1']='Примечание (необязательно)'
for r in range(2,OEND+1):os.cell(r,7,os.cell(r,7).value.replace(f'D{r}<>"Подтверждено",','').replace(f'{mget(f"B{r}",11)}<>"Подтверждено"','FALSE'))
os.data_validations.dataValidation=[v for v in os.data_validations.dataValidation if not str(v.sqref).startswith('D2:')]
# Extend issue indicators and explicit draft status.
check=wb['Проверки'];extra=[('Ошибки зон','=COUNTIF(Зоны!G2:G201,"Ошибка*")','Исправьте зоны.'),('Ошибки рейки / плинтуса','=COUNTIF(\'Рейка и плинтус\'!G2:G1001,"Ошибка*")','Проверьте длины, количество и зону.'),('Нормы не подтверждены','=IF(Параметры!B100="Да",0,1)','До подтверждения норм расчёт черновой.'),('Нулевые / неверные нормы','=IF(OR(STUDS_CART<=0,PLINTH_STEP<=0,TAPE_ROLL<=0,SEAL_M<=0,SHPON_L<=0,STEP_CONC<=0,STEP_MET<=0,STEP_VIB<=0,CONTOUR<0),1,0)','Делители и шаги должны быть положительными.'),('Зоны за пределом формы','=COUNTA(Зоны!A:A)-COUNTA(Зоны!A1:A201)','Не более200 зон.'),('Линейные элементы за пределом формы','=COUNTA(\'Рейка и плинтус\'!A:A)-COUNTA(\'Рейка и плинтус\'!A1:A1001)','Не более1000строк.')]
for row in extra:check.append(row)
wb['Итог']['B4']='=IF(Проверки!B12=0,"ПУСТАЯ ФОРМА",IF(SUM(Проверки!B2:B4,Проверки!B6:B11,Проверки!B13:B18)>0,"ЕСТЬ ЗАМЕЧАНИЯ / ЧЕРНОВИК","КОЛИЧЕСТВА ПРОВЕРЕНЫ"))'
wb['Инструкция'].delete_rows(1,wb['Инструкция'].max_row)
newinstructions=[('TSCalc Manual 0.2','Ручная геометрия, автоматические расходы материалов. DWG будет подключён после калькулятора.'),('1. Зоны','Введите ID, основание, площадь, виброизоляцию Да/Нет и периметр акустического шва (0 без него).'),('2. Параметры','Сохранены настройки старого Excel. Проверьте нормы, шаги,сечения,запасы,делители и наружный контур щитового поля. Нормы подтвердите один раз в B100 после проверки по проекту/паспортам.'),('3. Лаги','Введите ID, материал ЛАГ-001…004, длину одного физического куска в мм, количество одинаковых кусков, основание,узел,зонуID и физическое число стыков(0еслинет). Узел ДЕР-01: вручнуюточекнакусок; БЕТ/МЕТ/ВИБ: точкипо шагу каждойдетали+1. Правилообоихконцов требуетпроектногоподтверждения.'),('4. Щиты','ID шести допустимых сочетаний,целое количество,зонаID. Дополнительного подтверждениякаждойстрокине требуется. Полныещиты,раскройотдельно.'),('5. Рейка и плинтус','Введите фактические отрезки: вид,длинаодного в мм,количество,зона. Две нарисованные сторонырейки не удваиваются. Плинтус и акустическийпериметрраздельны.'),('6. Автоматические материалы','Восстановлены все31категориистарогоExcel: основание,брус,пластины,крепёж,войлок,виброизоляция,щиты,шпонка,клей,плинтус,рейка,краска. Вместо3строкщитов6вариантовкаталога. Прочие — толькодополнительныепозиции,не заменяютавторасход.'),('7. Цены и закупка','На листеМатериалы дляавтопозиций(начинаястроки30) задайтеценузаединицурасчёта,кратностьизапас. Запасодинраз,объединениедоокругления. Пустаяцена — неполнаястоимость. Монтаж/доставка/НДСотдельноне начисляются.'),('8. Проверки','Замечаниявлияютнаготовность. Количестваприошибкахчастичные. Спецификация — фильтрнепустыхID. Нормыизархива неявляютсяутверждёнными. 2000строклаг/1000щитов/1000линейных/200зон.'),('9. Ограничения','Войлоквэтойверсии — по всейдлине введённыхлаг: отключение/покрытиепроверитьвнастройках. Стандартныещиты2400×1200×60,2400×1200×70,2000×1000×60; ангарскаясосна/лиственница. Стыкивручную,пересеченияярусоввручную. Защитаотискажениякаталогапроверяетсяревью.')]
for i,(title,body) in enumerate(newinstructions,1):
 wb['Инструкция'].cell(i,2,title+'\n'+body);wb['Инструкция'].cell(i,2).alignment=Alignment(wrap_text=True,vertical='top');wb['Инструкция'].row_dimensions[i].height=90
## Position-level geometry owns surface calculations used by automatic coatings.
gs=init('Вычисления',['Вид','ID позиции','Зона ID','Материал ID','Ед. из базы','Длина, мм','Ширина, мм','Высота / толщина, мм','Количество','Верх щита, м²','Низ щита, м²','4 продольные / торцы щита, м²','2 торца лаги, м²','Полная поверхность, м²','Объём чисто, м³','Поверхность для покрытия, м²'],[19,20,20,20,20,20,20,25,20,23,23,38,25,29,25,36])
for r in range(2,LEND+1):
 for c,form in [(1,'"Лага"'),(2,f'Лаги!A{r}'),(3,f'Лаги!H{r}'),(4,f'Лаги!B{r}'),(5,mget(f'D{r}',4)),(6,f'Лаги!C{r}'),(7,f'Лаги!J{r}'),(8,f'Лаги!K{r}'),(9,f'Лаги!D{r}')]:gs.cell(r,c,f'=IF(Лаги!U{r}="ОК",{form},"")')
 gs.cell(r,10,0);gs.cell(r,11,0);gs.cell(r,12,f'=IF(B{r}="",0,F{r}*2*(G{r}+H{r})*I{r}/1000000)');gs.cell(r,13,f'=IF(B{r}="",0,2*G{r}*H{r}*I{r}/1000000)');gs.cell(r,14,f'=L{r}+M{r}');gs.cell(r,15,f'=IF(B{r}="",0,F{r}*G{r}*H{r}*I{r}/1000000000)');gs.cell(r,16,f'=L{r}')
 lags.cell(r,13,f'=Вычисления!O{r}')
 lags.cell(r,23).fill=PatternFill('solid',fgColor=yellow)
lags['W1']='Войлок Да/Нет';lags.column_dimensions['W'].width=20;dv(lags,'list','"Да,Нет"',f'W2:W{LEND}')
for pr in range(2,PEND+1):
 r=LEND+pr-1
 for c,form in [(1,'"Щит"'),(2,f'Щиты!A{pr}'),(3,f'Щиты!D{pr}'),(4,f'Щиты!B{pr}'),(5,mget(f'D{r}',4)),(6,f'Щиты!F{pr}'),(7,f'Щиты!G{pr}'),(8,f'Щиты!H{pr}'),(9,f'Щиты!C{pr}')]:gs.cell(r,c,f'=IF(Щиты!K{pr}="ОК",{form},"")')
 gs.cell(r,10,f'=IF(B{r}="",0,F{r}*G{r}*I{r}/1000000)');gs.cell(r,11,f'=J{r}');gs.cell(r,12,f'=IF(B{r}="",0,2*(F{r}+G{r})*H{r}*I{r}/1000000)');gs.cell(r,13,0);gs.cell(r,14,f'=J{r}+K{r}+L{r}');gs.cell(r,15,f'=IF(B{r}="",0,F{r}*G{r}*H{r}*I{r}/1000000000)');gs.cell(r,16,f'=K{r}+IF(EDGES="Да",L{r},0)')
 panels.cell(pr,9,f'=Вычисления!J{r}');panels.cell(pr,10,f'=Вычисления!O{r}')
finish(gs,LEND+PEND-1,[],'PositionGeometry')
ps['B86']=f'=SUM(Вычисления!P2:P{LEND})'
ps['B87']=f'=SUMPRODUCT(Вычисления!F2:F{LEND},Вычисления!G2:G{LEND},Вычисления!I2:I{LEND},--(Лаги!W2:W{LEND}="Да"))/1000000'
ps['B94']=f'=(SUMPRODUCT(Вычисления!F{LEND+1}:F{LEND+PEND-1},Вычисления!I{LEND+1}:I{LEND+PEND-1})+SUMPRODUCT(Вычисления!G{LEND+1}:G{LEND+PEND-1},Вычисления!I{LEND+1}:I{LEND+PEND-1}))*2/1000'
ps['B95']=f'=SUM(Вычисления!L{LEND+1}:L{LEND+PEND-1})'
ps['A104']='Окраска лаг: стороны';ps['B104']='4 продольные';ps['D104']='Как в старом Excel: 4продольные, без торцов. Полные6граней показаны отдельно на Вычисления.'
# Strict six-product standard catalogue: IDs cannot conceal altered dimensions.
for r in range(2,PEND+1):
 f=panels.cell(r,11).value
 allowed=[]
 for cr,(l,b,h) in enumerate([(2400,1200,60),(2400,1200,60),(2400,1200,70),(2400,1200,70),(2000,1000,60),(2000,1000,60)],2):allowed.append(f'AND(B{r}=Материалы!A{cr},F{r}={l},G{r}={b},H{r}={h},Материалы!C{cr}="{catalog[cr-2][2]}")')
 f=f.replace('"Ошибка: каталог щитов"', '"Ошибка: каталог щитов"')
 f=f.replace(f'{mget(f"B{r}",4)}<>"шт",',f'{mget(f"B{r}",4)}<>"шт",NOT(OR({",".join(allowed)})),')
 panels.cell(r,11,f)
# Clear misleading legacy instructions and preserve a visible geometric surface convention.
wb['Инструкция']['B10']='10. Вычисления\nПозиционно: щиты — верх,низ,торцы,полнаяповерхность; лаги —4продольные,2торца,полные6граней. Для автоматическойокраскилагприняты4продольныекаквстаромExcel. ДлявойлокавведитеДа/Нет в W листаЛаги. Стандартнаябазасодержит300строк; добавляйтематериалвпустуюстрокувнутритаблицы,выбирайтеIDввводах;единицыподтягиваютсяавтоматически.'
# Readable final user instructions supersede the incremental implementation notes above.
clean=[
('TSCalc Manual 0.2','Ручной ввод геометрии и зон. Материалы и крепёж рассчитываются автоматически. Чтение DWG будет подключено после завершения калькулятора.'),
('1. Зоны','Введите ID зоны, основание, площадь в м², наличие виброизоляции (Да/Нет) и периметр акустического шва в метрах. Если шва нет, введите 0. Зоны не должны перекрываться по площади.'),
('2. Параметры','Проверьте исторические настройки старого Excel: сечения, шаги, нормы расхода, запасы и длины упаковок. Наружный контур щитового поля задаётся отдельно в B9. После сверки с проектом и паспортами подтвердите нормы один раз в B100. До этого весь выпуск — черновой.'),
('3. Лаги','Одна строка означает одинаковые физические куски. Введите ID, материал ЛАГ-001…004, длину одного куска в мм, количество, основание, узел, зону, физическое число стыков (0, если нет) и войлок Да/Нет. Для дерева задайте точки крепления на кусок вручную. Бетон, металл и виброузел используют настройку шага.'),
('4. Узлы крепления','БЕТ-01 — бетон; МЕТ-01 — металл; ДЕР-01 — дерево; ВИБ-01 — виброопоры. По шагу: ceil(длина одного куска / шаг) + 1, затем умножение на количество кусков. Это правило двух концов применяется только после подтверждения узла. Граница зоны сама по себе не создаёт стык.'),
('5. Щиты','Выберите ID из базы, введите целое количество и ID зоны. В базе только три размера: 2400×1200×60, 2400×1200×70 и 2000×1000×60 мм. Для каждого доступны ангарская сосна и лиственница. Вводятся полные закупаемые изделия. Раскрой обрезанных деталей в эту версию не входит.'),
('6. Рейка и плинтус','Введите вид, длину одного отрезка в мм, количество и зону. Учитывайте фактическую длину дубовой рейки: две стороны шва не удваиваются повторно. Плинтус и периметр акустического шва вводятся раздельно.'),
('7. Автоматические расходы','Сохранены все 31 категории старого Excel. Расходы основания, бруса, пластин, крепежа, войлока, звукоизоляции, шпонки, клея, реек, плинтуса и краски рассчитываются из геометрии и настроек. Три старые строки щитов заменены шестью вариантами базы. Лист Прочие служит только для дополнительных позиций.'),
('8. Материалы и цены','Добавляйте материалы в пустые строки базы внутри её границ. Выбирайте ID в выпадающих списках. Единица измерения подтягивается из базы. Для автоматических позиций с 30-й строки задайте цену за единицу расчёта и кратность. Запас применяется один раз, закупка округляется после объединения материала по объекту.'),
('9. Вычисления','На каждую позицию показаны ID, зона, материал и единица из базы. Щиты: верх, низ, торцы и полная поверхность. Лаги: четыре продольные грани, два торца и полная поверхность. Для автоматической окраски лаг приняты четыре продольные грани, как в старом Excel; торцы показаны отдельно.'),
('10. Проверка и выпуск','Исправьте все замечания на листе Проверки. Пустая цена означает неполную стоимость. При ошибках количества частичные. На листе Спецификация отфильтруйте пустые ID перед печатью. Монтаж, доставка, накладные, прибыль и отдельное начисление НДС в этой версии не добавляются.'),
('Границы формы','2000 строк лаг, 1000 строк щитов, 1000 линейных элементов, 200 зон, 300 материалов, 100 узлов. Не вводите данные за границей формы. Для расширения требуется проверить формулы. Жёлтые ячейки и синий шрифт — ввод; серые — вычисления.')]
ins=wb['Инструкция'];ins.delete_rows(1,ins.max_row)
for i,(title,body) in enumerate(clean,1):
 ins.cell(i,2,title+'\n'+body);ins.cell(i,2).alignment=Alignment(wrap_text=True,vertical='top');ins.cell(i,2).font=Font(name='Calibri',size=12,bold=i==1,color=navy);ins.row_dimensions[i].height=90
ins.print_area=f'B1:B{len(clean)}'
# Final narrow input-integrity guards.
for r in range(2,LEND+1):
 f=lags.cell(r,21).value
 f=f.replace(f'COUNTIF(Зоны!$A$2:$A$201,H{r})<>1,',f'COUNTIF(Зоны!$A$2:$A$201,H{r})<>1,IFERROR(VLOOKUP(H{r},Зоны!$A$2:$G$201,2,FALSE),"")<>E{r},IF(F{r}="ВИБ-01",IFERROR(VLOOKUP(H{r},Зоны!$A$2:$G$201,4,FALSE),"")<>"Да",IFERROR(VLOOKUP(H{r},Зоны!$A$2:$G$201,4,FALSE),"")="Да"),')
 lags.cell(r,21,f)
for r in range(2,MEND+1):
 f=cs.cell(r,13).value.replace('"Ошибка: запас / кратность"','"Ошибка: единица / запас / кратность"')
 f=f.replace(f'NOT(ISNUMBER(E{r})),',f'AND(C{r}<>"шт",C{r}<>"м",C{r}<>"м²",C{r}<>"м³",C{r}<>"кг",C{r}<>"л",C{r}<>"рул"),AND(C{r}="шт",MOD(H{r},1)<>0),NOT(ISNUMBER(E{r})),')
 cs.cell(r,13,f)
for v in ms.data_validations.dataValidation:
 if str(v.sqref).startswith('D2:'):v.formula1='"шт,м,м²,м³,кг,л,рул"'
for rr in range(2,check.max_row+1):
 if check.cell(rr,1).value=='Нулевые / неверные нормы':
  check.cell(rr,2,'=IFERROR(IF(OR(STUDS_CART<=0,PLINTH_STEP<=0,TAPE_ROLL<=0,SEAL_M<=0,SHPON_L<=0,STEP_CONC<=0,STEP_MET<=0,STEP_VIB<=0,CONTOUR<0,AND(N_PANELS>0,OR(CONTOUR<=0,CONTOUR>PER_PANELS))),1,0),1)')
for rr in range(30,30+len(auto)):
 code=ms.cell(rr,1).value
 if code and code.startswith('БРУС'):
  br=int(code[-1]);ms.cell(rr,3,f'=Материалы!C{7+br}')
# Detect partially filled records beyond the explicit capacities, not just their IDs.
for rr,sheet,last,lastletter in [(7,'Лаги',LEND,'I'),(8,'Щиты',PEND,'E'),(9,'Прочие',OEND,'E'),(10,'Материалы',MEND,'L'),(11,'Узлы',NEND,'L'),(17,'Зоны',201,'F'),(18,'Рейка и плинтус',1001,'E')]:
 q=f"'{sheet}'";check.cell(rr,2,f'=COUNTA({q}!A:{lastletter})-COUNTA({q}!A1:{lastletter}{last})')
for rr in range(3,11):
 wb['Итог'].cell(rr,2).fill=PatternFill('solid',fgColor='F2F2F2');wb['Итог'].cell(rr,2).font=Font(name='Calibri',size=11,color='222222')
# Sanity-check formula syntax balance before export (not a substitute for native recalculation).
lags['X1']='Высота для бетонного узла';lags['Y1']='Высота для металлического узла'
for r in range(2,LEND+1):
 lags.cell(r,24,f'=IF(AND(U{r}="ОК",E{r}="Бетон",F{r}<>"ВИБ-01"),K{r},0)');lags.cell(r,25,f'=IF(AND(U{r}="ОК",E{r}="Металл",F{r}<>"ВИБ-01"),K{r},0)')
lags.column_dimensions['X'].hidden=True;lags.column_dimensions['Y'].hidden=True
ps['B90']=f'=MAX(Лаги!X2:X{LEND})';ps['B91']=f'=MAX(Лаги!Y2:Y{LEND})'
check.append(('Недостаточная длина шпильки','=IFERROR(IF(AND(P_CONC>0,STUD_L<HMAX_CONC+PAD_T+EMBED+20),1,0),1)','Сквозной узел: высота лаги + подкладка + заделка + 20 мм гайка/шайба. Для уголка требуется отдельная проверка.'))
check.append(('Недостаточная длина болта','=IFERROR(IF(AND(P_MET>0,BOLT_L<HMAX_MET+PAD_T+FLANGE_T+20),1,0),1)','Высота лаги + подкладка + полка металла + 20 мм гайки/шайбы.'))
check.append(('Недостаточная заделка самореза щита','=IFERROR(IF(AND(N_PANELS>0,SCREW_L-(T_MAX-PLUG_D)-FELT_T<MIN_EMB),1,0),1)','Длина самореза − (толщина щита − утопление пробки) − войлок; сравнить с минимальной заделкой.'))
wb['Итог']['B4']='=IF(Проверки!B12=0,"ПУСТАЯ ФОРМА",IF(SUM(Проверки!B2:B4,Проверки!B6:B11,Проверки!B13:B21)>0,"ЕСТЬ ЗАМЕЧАНИЯ / ЧЕРНОВИК","КОЛИЧЕСТВА ПРОВЕРЕНЫ"))'
finish(check,check.max_row,[])
for sn,col,end in [('Лаги',21,LEND),('Щиты',11,PEND),('Прочие',7,OEND),('Зоны',7,201),('Рейка и плинтус',7,1001),('Расчёт',13,MEND)]:
 for rownum in range(2,end+1):
  c=wb[sn].cell(rownum,col);c.value='=IFERROR('+c.value[1:]+',"Ошибка: неверный тип данных")'
for sheet in wb:
 for row in sheet:
  for cell in row:
   if cell.data_type=='f':
    expr=cell.value; quoted=False;level=0
    for ch in expr:
     if ch=='"':quoted=not quoted
     elif not quoted and ch=='(':level+=1
     elif not quoted and ch==')':level-=1
    if level!=0:raise ValueError(f'Unbalanced formula {sheet.title}!{cell.coordinate}: {expr}')
wb.active=0
OUT.parent.mkdir(parents=True,exist_ok=True);wb.save(OUT)
print(OUT)
print('SHA256',hashlib.sha256(OUT.read_bytes()).hexdigest())
