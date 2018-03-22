__precompile__()

module EzExcel

using Missings
using EzXML, ZipFile


# import for use
using ZipFile: ReadableFile, Reader

include("parse.jl")
include("index.jl")
include("format.jl")

include("workbook.jl")
include("worksheet.jl")
include("cell.jl")


export WorkBook, 
        sheetnames, sheetname,
        peeloff

end
