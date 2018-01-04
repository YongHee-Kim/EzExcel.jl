# Teaked DataFrames Index for SharedStrings and SheetName
# https://github.com/JuliaData/DataFrames.jl/blob/master/src/other/index.jl
struct Index
    lookup::Dict{String, Int}
    names::Vector{String}
end

function Index(names::Vector{String})
    if unique(names) != names
        error("Excel do not allow duplicated string for SharedStrings and SheetName")
    end
    lookup = Dict{String, Int}(zip(names, 1:length(names)))
    Index(lookup, names)
end
