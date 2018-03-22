# Teaked DataFrames Index for SharedStrings and SheetName
# https://github.com/JuliaData/DataFrames.jl/blob/master/src/other/index.jl
struct Index
    lookup::Dict{String, Int}
    names::Tuple
end

function Index(names::Vector{String})
    if unique(names) != names
        error("Excel do not allow duplicated string for SheetName")
    end
    lookup = Dict{String, Int}(zip(names, 1:length(names)))
    Index(lookup, Tuple(names))
end
Index() = Index(Dict{String, Int}(), Tuple{}())

Base.length(x::Index) = length(x.names)
Base.names(x::Index) = copy(x.names)

Base.getindex(x::Index, idx::String) = x.lookup[idx]
Base.getindex(x::Index, idx::Real) = x.names[idx]
