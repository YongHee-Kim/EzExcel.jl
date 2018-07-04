# EzExcel.jl
[![License](http://img.shields.io/badge/license-MIT-brightgreen.svg?style=flat)](LICENSE)

EzExcel is Excel file reader in pure Julia.
This is not actively being maintained, I recommend anyone come across here to use [XLSX](https://github.com/felipenoris/XLSX.jl)


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

# You can access `WorkSheet` via index or sheet name

sheet1 = wb[1] 
sheet3 = wb["Second Sheet"]

