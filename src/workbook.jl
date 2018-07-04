
"""
    `WorkBook` contains number of `WorkSheet` objects
    `WorkSheet` can be accessed by index number or name of WorkSheet

    **Arguments**
"""
struct WorkBook
    path::String
    custom_number_format::Dict
    xf_index_to_number_format::Tuple
    shared_strings::Tuple

    sheets::Vector
    sheet_index::Index
end

# Constructor
function WorkBook(path::AbstractString)
    if !endswith(path, ".xlsx")
        throw(DomainError("EzExcel is only compatible with '.xlsx' format"))
    end
    file = ZipFile.Reader(path)
    # Shared informations in WorkBook
    style_xml = get_xmlroot(file, "styles.xml")
    workbook_xml = get_xmlroot(file, "workbook.xml")
    sharedstrings_xml = get_xmlroot(file, "sharedstrings.xml")

    custom_number_format = extract_numFmts(style_xml)
    xf_index_to_number_format = extract_cellXfs(style_xml)
    s_strings = extract_sharedstrings(sharedstrings_xml)

    # initialize WorkBook for parenting WorkSheet
    sheet_names = extract_sheetnames(workbook_xml)
    wb = WorkBook(path, custom_number_format, xf_index_to_number_format, s_strings,
             Array{WorkSheet}(length(sheet_names)), Index(sheet_names))
    WorkBook(wb, file)
end
function WorkBook(wb::WorkBook, file)
    for (i, name) in enumerate(sheetname(wb))
        p = "xl\\worksheets\\sheet$i.xml"
        wb.sheets[i] = WorkSheet(wb, name, get_xmlroot(file, p))
    end
    close(file)
    wb
end


function get_xmlroot(file, el)
    x = filter(x -> endswith(lowercase(normpath(x.name)), el), file.files)
    isempty(x) ? nothing : root(parsexml(x[1]))
end
"""
    extract_cellXfs(xml)

order of values in `cellXfs` is cell style index number
설명 수정 필요 stackoverflow 답변 링크?
"""
function extract_cellXfs(xml)
    xf_node = Iterators.filter(x -> nodename(x) == "cellXfs", eachnode(xml))
    xf_node = collect(xf_node)[1] |> nodes

    xf_datas = filter(iselement, xf_node)    

    fmt_ids = zeros(UInt16, length(xf_datas))
    for (i, x) in enumerate(xf_datas)
        if haskey(x, "numFmtId") 
            fmt_ids[i] = parse(UInt16, x["numFmtId"])
        end
    end
    Tuple(fmt_ids)
end

"""
    extract_numFmts

find datatype of cell value based on format code info.
some formate codes ared predefined in `ECMA_NUMBER_FORMAT`, but others are not predefined and
defined within indivisual Excel File

설명 수정 필요...

reference:[ECMA-376 standard](http://www.ecma-international.org/publications/standards/Ecma-376.htm)
           5th Edition Part1 - 18.8.30 numFmt (Number Format)
"""
function extract_numFmts(xml)
    d = Dict{UInt16,Tuple}()
    fmt_node = Iterators.filter(x -> nodename(x) == "numFmts", eachnode(xml))

    if !isempty(fmt_node)
        fmt_datas = collect(fmt_node)[1] |> nodes

        for el in filter(iselement, fmt_datas)
            fmtid = parse(Int, el["numFmtId"])
            d[fmtid] = (guess_datatype(el["formatCode"]), el["formatCode"])
        end
    end
    d
end

function extract_sharedstrings(xml)
    if isa(xml, Void)
        ()
    else
        x = elements(xml) |> x -> vcat(elements.(x)...)
        x = filter(el -> nodename(el) == "t", x)
        Tuple(nodecontent.(x))
    end
end

function extract_sheetnames(xml)
    sheets_node = filter(x -> nodename(x) .== "sheets", elements(xml))
    sheets_node = elements(sheets_node[1])

    k = haskey(sheets_node[1], "r:id") ? "r:id" : 
        haskey(sheets_node[1], "d3p1:id") ? "d3p1:id" : 
        error("sheet index is not readable")
    
    sheet_cnt = length(sheets_node)
    rids = Vector{Int}(sheet_cnt)
    sheet_names = Vector{String}(sheet_cnt)
    
    for (i, el) in enumerate(sheets_node)
        rids[i] = split(el[k], "rId")[2] |> x -> parse(Int, x)
        sheet_names[i] = el["name"]
    end

    return sheet_names[sortperm(rids)]
end


##############################################################################
##
## Basic properties of a WorkBook
##
##############################################################################

function Base.getindex(wb::WorkBook, inds...)
    for (i, j) in enumerate(inds...)
        ws = wb.sheets[j]
        if isa(ws, WorkSheet{EzXML.Node})
            wb.sheets[j] = WorkSheet(ws)
        end
    end
    getindex(wb.sheets, inds...)
end
Base.getindex(wb::WorkBook, key::String) = getindex(wb, wb.sheet_index[key])

Base.length(wb::WorkBook) = length(wb.sheets)
# Base.start(wb::WorkBook) = start(wb.sheets)
# Base.next(wb::WorkBook, i) = next(wb.sheets, i)
# Base.done(wb::WorkBook, i) = done(wb.sheets, i)
Base.endof(wb::WorkBook) = endof(wb.sheets)


# Accessors for WorkBook
"""
A list of all sheets in the WorkBook
"""
sheetname(wb::WorkBook) = wb.sheet_index.names
sheets(wb::WorkBook) = wb[1:end]
