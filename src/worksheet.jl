abstract type ExcelSheet end
"""
    Contains the data for one worksheet.
"""
struct WorkSheet
    name::AbstractString
    header::Vector{Symbol}
    data
end
function WorkSheet(xml)

end




"""
    col_number(col::AbstractString)
convert Excel column name(A ~ XFD) to column number(1~16384)
"""
function col_number(col::AbstractString)
    cl = uppercase(col)
    r = 0
    for c in cl
        r = (r * 26) + (c - 'A' + 1)
    end
    return r
end


function convert_ref_to_sheet_row_col(range::AbstractString)
    r = r"('?[^']+'?|[^!]+)!([A-Za-z]*)(\d*)(:([A-Za-z]*)(\d*))?"
    m = match(r, range)
    m == nothing && error("Invalid Excel range specified.")

    sheetname = String(m.captures[1])
    startrow = parse(Int, m.captures[3])
    startcol = col_number(m.captures[2])
    if m.captures[4] == nothing
        endrow = startrow
        endcol = startcol
    else
        endrow = parse(Int, m.captures[6])
        endcol = col_number(m.captures[5])
    end
    if (startrow > endrow ) || (startcol > endcol)
        error("Please provide rectangular region from top left to bottom right corner")
    end
    return sheetname, startrow, startcol, endrow, endcol
end
