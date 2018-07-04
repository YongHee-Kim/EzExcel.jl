__precompile__()

module EzExcel

@static if VERSION < v"0.7.0-DEV.2005"
    using Missings
end
using EzXML
using ZipFile
using ZipFile: ReadableFile, Reader

include("parse.jl")
include("index.jl")
include("format.jl")

include("workbook.jl")
include("worksheet.jl")
include("cell.jl")
include("show.jl")


export WorkBook, WorkSheet, Cell,
       sheetname, sheets,
       size, length

end
