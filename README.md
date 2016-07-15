# Excel2CSV_NPOI
use NPOI to rewrite the Excel2CSV app

## 2016-07-14
it seems that color problem can be solved,but not in a good way.  
and before using any object you must check it`s not null.  
really annying...  
but I check the output excel file,not cell or row lost..

## 2016-07-15
OK~I find a lib(not free,a cracked version) help me convert xslx to csv.But there is more problem with NPOI waiting for me.  
NPOI and sorting or filter the sheet.The only way to sovle this problem is copy every cell to a datatable and then copy them back.......  
I know this process is very fast,because it is in memory.  
But I already use a not free lib calls Gembox,why I use NPOI for import data from xlsx......  
Maybe there is some efficient problem,but who knows.  
Because Gembox cost a lot,not a lot discuss on net...  
Anyway,my priority is compatibility... 
So I will try Gembox~  