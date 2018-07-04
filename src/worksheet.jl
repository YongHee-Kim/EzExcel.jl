"""
    WorkSheet
Contains the data for one worksheet.
"""
struct WorkSheet{T}
    book::WorkBook
    name::AbstractString
    data::T
end
function WorkSheet(ws::WorkSheet, preserve_cell = false)
    function parse_row(row)
        v = Array{Union{Cell, Missing}}(1, data_size[2]) .= missing
        for cell in filter(iselement, nodes(row))
            m = match(r"([A-Za-z]*)", cell["r"])
            col_num = col_number(m.captures[1]) - dimension[2] + 1
            v[1, col_num] = Cell(ws.book, cell)
        end
        v
    end
    xml = ws.data

    dimension = extract_dimension(xml)
    data_size = (dimension[3] - dimension[1] +1, dimension[4] - dimension[2] +1)

    # reading by row
    rowdata_node = Iterators.filter(x -> nodename(x) == "sheetData", eachnode(xml))
    row_datas = collect(rowdata_node)[1] |> nodes

    data = Array{Union{Missing, Cell}}(data_size) .= missing
    for x in filter(iselement, row_datas)
        row_num = parse(Int, x["r"]) - dimension[1] +1
        data[row_num, :] = parse_row(x)
    end

    if preserve_cell
        WorkSheet(ws.book, ws.name, data)
    else
        WorkSheet(ws.book, ws.name, _peeloff(data))
    end
end

function parse_xlsx_row(row, data_size)
    v = Array{Union{Cell, Missing}}(1, data_size[2]) .= missing
    for cell in filter(iselement, nodes(row))
        m = match(r"([A-Za-z]*)", cell["r"])
        col_num = col_number(m.captures[1]) - dimension[2] + 1
        v[1, col_num] = Cell(ws.book, cell)
    end
    v
end
function extract_dimension(xml)
    dim = Iterators.filter(x -> nodename(x) == "dimension", eachnode(xml)) |> collect
    if isempty(dim)
        throw(ParseError("Dimension information is missing, open and resaving .xlsx from Office will solve this issue"))
    end
    x = dim[1]["ref"]
    !contains(x, ":") && (x = ("A1:" * x))

    convert_ref_to_row_col(x)
end


"""
    col_number(col::AbstractString)
convert Excel column name(A ~ XFD) to column number(1~16384)
"""
function col_number(col::AbstractString)
    r = 0
    for c in uppercase(col)
        r = (r * 26) + (c - 'A' + 1)
    end
    return r
end

"""
    convert_ref_to_row_col(range::AbstractString)

converts 'A1' cell notation to 'R1C1' notation
"""
function convert_ref_to_row_col(range::AbstractString)
    # r = r"('?[^']+'?|[^!]+)!([A-Za-z]*)(\d*)(:([A-Za-z]*)(\d*))?"
    r = r"([A-Za-z]*)(\d*)"
    m = match.(r, split(range, ":"))
    m[1] == nothing && error("Invalid Excel range specified.")

    startrow = parse(Int, m[1].captures[2])
    startcol = col_number(m[1].captures[1])

    if length(m) == 1
        endrow = startrow
        endcol = startcol
    else
        endrow = parse(Int, m[2].captures[2])
        endcol = col_number(m[2].captures[1])
    end
    if (startrow > endrow ) || (startcol > endcol)
        error("Please provide rectangular region from top left to bottom right corner")
    end
    return startrow, startcol, endrow, endcol
end
function convert_ref_to_sheet_row_col(range::AbstractString)
    s = split(range, "!")
    return s[1], convert_ref_to_row_col(s[2])...
end

##############################################################################
##
## Basic properties of a WorkSheet
##
##############################################################################
Base.size(a::WorkSheet) = size(a.data)
Base.size(a::WorkSheet, d) = size(a.data, d)
Base.length(a::WorkSheet) = length(a.data)

Base.getindex(ws::WorkSheet, inds...) = getindex(ws.data, inds...)
function Base.getindex(ws::WorkSheet, range::AbstractString) 
    rg = convert_ref_to_row_col(range)
    getindex(ws.data, rg[1]:rg[3], rg[2]:rg[4])
end

sheetname(ws::WorkSheet) = ws.name