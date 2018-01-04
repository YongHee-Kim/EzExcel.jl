# convert ExcelErrorText to Excel internal code for error cells.
struct ExcelErrorCell
    errorcode::UInt8
end
ExcelErrorCell(e::AbstractString) = ExcelErrorCell(EXCEL_ERROR_CODE[e])


function show(io::IO, o::ExcelErrorCell)
    print(io, o.errorcode)
end

const EXCEL_ERROR_CODE = Dict(
    "#NULL!" => 0x00,       # Intersection of two cell ranges is empty
    "#DIV/0!"=> 0x07,       # Division by zero
    "#VALUE!"=> 0x0F,       # Wrong type of operand
    "#REF!"  => 0x17,       # Illegal or deleted cell reference
    "#NAME?" => 0x1D,       # Wrong function or range name
    "#NUM!"  => 0x24,       # Value range overflow
    "#N/A"   => 0x2A)       # Argument or function not available
