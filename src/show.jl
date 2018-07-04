function Base.show(io::IO, wb::WorkBook)
    print(io, basename(wb.path), " (with ", length(wb), " WorkSheet)") 
end

function Base.show(io::IO, ws::WorkSheet)
    print("WorkSheet[$(ws.name)] ")
    display(ws.data)
end

function Base.show(io::IO, x::Cell)
    print(io, x.v)
end



