EzXML.parsexml(x::ReadableFile) = parsexml(readstring(x))
EzXML.parsexml(x::Reader) = parsexml(readstring(x))

"""
    Contents of a "workbook"
"""
struct WorkBook
    file::ZipFile.Reader
    shared_strings::Index
    sheets::Index
    children::Tuple
    parsed_children::Vector{Union{Missing, WorkSheet}}
end
function WorkBook(filename::AbstractString)
    file = ZipFile.Reader(filename)
    s_strings = shared_strings(file)
    sheets = sheetnames(file)

    children = begin
        xml_files = broadcast(i -> "xl\\worksheets\\sheet$i.xml", eachindex(sheets))
        ind = indexin(xml_files, broadcast(x -> normpath(x.name), file.files))
        file.files[ind]
    end
    parsed_children = Union{Missing, WorkSheet}[fill(missing, length(sheets))...]
    WorkBook(file, Index(s_strings), Index(sheets), Tuple(children), parsed_children)
end


function Base.show(io::IO, wb::WorkBook)
    print(io, wb.sheets.lookup)
end


function getindex(wb::WorkBook, inds...)
    _valid = wb.parsed_children |> x -> broadcast(i -> isassigned(x, i), inds...)

    for i in eachindex(wb.parsed_children)

    end
    getindex(wb.parsed_children, inds...)
end
function getindex(wb::WorkBook, key)

end
getindex(wb::WorkBook, key::Symbol) = getindex(wb, string(key))

###
"""
A list of all sheets in the WorkBook
"""
function sheetnames(xlfile)
    xml_file = filter(x -> endswith(normpath(x.name), "xl\\workbook.xml"), xlfile.files)

    f = readstring(xml_file[1]) |> parsexml |> root

    sheets_node = filter(x -> nodename(x) .== "sheets", elements(f))
    sheets = Vector{String}(countelements(sheets_node[1]))
    for el in elements(sheets_node[1])
        sheets[parse(el["sheetId"])] = el["name"]
    end

    return sheets
end

function shared_strings(xlfile)
    xml_file = filter(x -> endswith(lowercase(x.name), "sharedstrings.xml"), xlfile.files)

    if !isempty(xml_file)
        f = readstring(xml_file[1])
        if !isempty(f)
            f = root(parsexml(f))

            return nodecontent.(collect(eachnode(f)))
        end
    end
    return nothing
end
