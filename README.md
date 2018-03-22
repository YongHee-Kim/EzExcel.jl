# EzExcel.jl

> EzExcel is not stable enought to be registered package, and I'm not planning to make it anytime soon. So, you are welcome to fork and develop separate package as you like.<br>
> EzExcel is heavily inspired by [ExcelReaders](https://github.com/davidanthoff/ExcelReaders.jl), [XLSXread](https://github.com/bbrunaud/XLSXread.jl/blob/master/src/XLSXread.jl) and [xlrd](https://github.com/python-excel/xlrd). Many lines of codes came from those projects

Excel file reader for Julia with minimal dependency on [EzXML](https://github.com/bicycle1885/EzXML.jl) and [ZipFile](https://github.com/fhs/ZipFile.jl). EzExcel is Only compatible with `XLSX` format that follows [ECMA-376 standard](http://www.ecma-international.org/publications/standards/Ecma-376.htm)


## Installation
``` Julia
Pkg.clone("https://github.com/YongHee-Kim/EzExcel.jl")
```
## Usage Example
```Julia
using EzExcel

path = Pkg.dir("EzExcel")

# Create WorkBook from XLSX file
wb = WorkBook("$path\\test\\datasets\\testdata.xlsx")

# Access `WorkSheet` via index or sheet name

sheet1 = wb[1] 
sheet3 = wb["Second Sheet"]

# `Cell` datas are stored in `data` field in `WorkSheet`
sheet1.data
sheet3.data

# you have to `peeloff` `Cell` to use loaded excel data 
using DataFrames

function consturct_dataframe(ws)
    data = ws.data
    header = Symbol.(peeloff(data[1, :]))
    table = map(i -> peeloff(data[2:end, i]), 1:size(data, 2))

    DataFrame(table, header)
end

df_sheet1 = consturct_dataframe(sheet1)
df_sheet3 = consturct_dataframe(sheet3)
```


