using EzExcel

using Base.Test


path = "c:\\Github\\EzExcel.jl\\test\\datasets"
test_files = readdir(path)

v = Any[]
for f in test_files
    println(f)
    push!(v, WorkBook("$path\\$f"))
end

# parse WorkSheet
for i in 1:length(v)
    println("parsing WorkSheet", v[i])
    try
        v[i][1:end]
    catch e
        println(e)
    end
end