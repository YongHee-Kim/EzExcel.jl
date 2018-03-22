"""
    Cell{T}

preserve 'fmtId' for Cell for future use
currently there is no functionality for fmtId after cell value is extracted
"""
struct Cell{T}
    numFmtId::UInt16
    v::T
end
Cell(id, v::T) where T = Cell{T}(id, v)

function Cell(wb::WorkBook, x::EzXML.Node)
    # println("cell data / ", x["r"])
    fmt_id, T = find_cell_datatype(wb, x)

    cell_values = filter(iselement, collect(eachnode(x)))
    # sometimes there can be no value, even through it has cell address
    value = if isempty(cell_values)
                missing
            else # Excel always store value `v` node at the end
                v = nodecontent.(cell_values)[end]
                v = get_cellvalue(T, wb, v)
            end
    
    Cell(fmt_id, value)
end

function Base.show(io::IO, x::Cell)
    print(io, x.v)
end

# fallback methods
Base.Symbol(x::Cell) = convert(Symbol, x)
Base.convert(::Type{Symbol}, x::Cell) = Symbol(x.v)

function find_cell_datatype(wb::WorkBook, x)
    fmt_id = haskey(x, "s") ? wb.xf_index_to_number_format[parse(Int, x["s"])+1] : 0
    T = if haskey(x, "t")
            ECMA_CELLTYPE[x["t"]]
        else
            fmt = haskey(ECMA_NUMBER_FORMAT, fmt_id) ? ECMA_NUMBER_FORMAT[fmt_id] :
                  wb.custom_number_format[fmt_id]
            fmt[1]
        end
    (fmt_id, T)
end

function get_cellvalue(::Type{Bool}, wb, x)::Bool
    if x == "1" || x == "0"
        x = parse(Int8, x)
    end
    Bool(x)
end
function get_cellvalue(::Type{DateTime}, wb, x)
    x
end
get_cellvalue(::Type{ExcelDateTime}, wb, x) = parse(ExcelDateTime, x)
get_cellvalue(::Type{Missing}, wb, x) = x
get_cellvalue(::Type{String}, wb, x) = String(x)
get_cellvalue(::Type{Float64}, wb, x) = parse(Float64, x)
get_cellvalue(::Type{T}, wb, x) where T<:Integer= parse(T, x)
function get_cellvalue(::Type{SharedString}, wb, x) 
    wb.shared_strings[parse(Int, x)+1]
end
get_cellvalue(::Type{GeneralFormat}, wb, x) = parse(GeneralFormat, x)
get_cellvalue(::Type{ScientificFormat}, wb, x) = parse(ScientificFormat, x)

"""
    peeloff(x::Vector{Union{Cell, Missing}})

peel off 'Cell{T}' and returns Array{Union{T...} with cell values
"""
function peeloff(datas::Vector{Union{Cell, Missing}})
    v = map(x -> ismissing(x) ? x : peeloff(x), datas)
    T = unique(typeof.(v))
    if length(T) < 3
        convert(Vector{Union{T...}}, v)
    else
        convert(Vector{Union{Missing, Any}}, v)
    end
end
peeloff(x::Cell) = x.v
