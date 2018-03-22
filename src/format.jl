"""
    ECMA_NUMBER_FORMAT

reference:[ECMA-376 standard](http://www.ecma-international.org/publications/standards/Ecma-376.htm)
           5th Edition Part1 - 18.8.30 numFmt (Number Format)

"""
const ECMA_NUMBER_FORMAT = Dict{UInt16, Tuple}(
   0 => (GeneralFormat,   "General"),
   1 => (Int,             "0"),
   2 => (Float64,         "0.00"),
   3 => (Float64,         "#,##0"),
   4 => (Float64,         "#,##0.00"),
   # not specified by ECMA-376 standard
   5 => (Float64,         "0.00"),
   6 => (Float64,         "0.00"),
   7 => (Float64,         "0.00"),
   8 => (Float64,         "0.00"),
   9 => (Float64,         "0%"),
  10 => (Float64,         "0.00%"),
  11 => (ScientificFormat,"0.00E+00"),
  12 => (Float64,         "# ?/?"),
  13 => (Float64,         "# ??/??"),
  14 => (ExcelDateTime,   "m/d/yy"),
  15 => (ExcelDateTime,   "d-mmm-yy"),
  16 => (ExcelDateTime,   "d-mmm"),
  17 => (ExcelDateTime,   "mmm-yy"),
  18 => (ExcelDateTime,   "h:mm AM/PM"),
  19 => (ExcelDateTime,   "h:mm:ss AM/PM"),
  20 => (ExcelDateTime,   "h:mm"),
  21 => (ExcelDateTime,   "h:mm:ss"),
  22 => (ExcelDateTime,   "m/d/yy h:mm"),
  27 => (ExcelDateTime,   "yyyy\"年\" mm\"月\" dd\"日\""),
  28 => (ExcelDateTime,   "mm-dd"),
  29 => (ExcelDateTime,   "mm-dd"),
  30 => (ExcelDateTime,   "mm-dd-yy"),
  31 => (ExcelDateTime,   "mm\"월\" dd\"일\""),
  32 => (ExcelDateTime,   "h\"시\" mm\"분\""),
  33 => (ExcelDateTime,   "h\"시\" mm\"분\" ss\"초\""),
  34 => (ExcelDateTime,   "yyyy-mm-dd"),
  35 => (ExcelDateTime,   "yyyy-mm-dd"),
  36 => (ExcelDateTime,   "mm\"月\" dd\"日\""),
  37 => (Float64,         "#,##0_);(#,##0)"),
  38 => (Float64,         "#,##0_);[Red](#,##0)"),
  39 => (Float64,         "#,##0.00_);(#,##0.00)"),
  40 => (Float64,         "#,##0.00_);[Red](#,##0.00)"),
  41 => (Float64,         "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)"),
  42 => (Float64,         "_(u0024* #,##0_);_(u0024* (#,##0);_(u0024* \"-\"_);_(@_)"),
  43 => (Float64,         "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)"),
  44 => (Float64,         "_(u0024* #,##0.00_);_(u0024* (#,##0.00);_(u0024* \"-\"??_);_(@_)"),
  45 => (ExcelDateTime,   "mm:ss"),
  46 => (ExcelDateTime,   "[h]:mm:ss"),
  47 => (ExcelDateTime,   "mm:ss.0"),
  48 => (ScientificFormat,"##0.0E+0"),
  49 => (String,          "@"),
  50 => (ExcelDateTime,   "mm\"月\" dd\"日\""),
  51 => (ExcelDateTime,   "mm-dd"),
  52 => (ExcelDateTime,   "yyyy-mm-dd"),
  53 => (ExcelDateTime,   "yyyy-mm-dd"),
  54 => (ExcelDateTime,   "mm-dd"),
  55 => (ExcelDateTime,   "yyyy-mm-dd"),
  56 => (ExcelDateTime,   "yyyy-mm-dd"),
  57 => (ExcelDateTime,   "mm\"月\" dd\"日\""),
  58 => (ExcelDateTime,   "mm-dd"))


"""
    ECMA_CELLTYPE
reference:[ECMA-376 standard](http://www.ecma-international.org/publications/standards/Ecma-376.htm)
           5th Edition Part1 - 18.18.11 ST_ST_CellType (Cell Type)
"""
const ECMA_CELLTYPE = Dict(
    "b"         => Bool,            # a Boolean
    "d"         => DateTime,        # a date in the ISO 8601 format
    "e"         => Missing,         # an error
    "inlineStr" => String,          # (inline) rich string
    "n"         => Float64,         # a Number
    "str"       => String,
    "s"         => SharedString)

"""

reference:[ECMA-376 standard](http://www.ecma-international.org/publications/standards/Ecma-376.htm)
        5th Edition Part1 - L.2.16.8 Error values
"""
 const ECMA_ERROR_VALUES = Dict(
     "#NULL!" => "Intersection of two cell ranges is empty",
     "#DIV/0!"=> "Division by zero",
     "#VALUE!"=> "Wrong type of operand",
     "#REF!"  => "Illegal or deleted cell reference",
     "#NAME?" => "Wrong function or range name",
     "#NUM!"  => "Value range overflow",
     "#N/A"   => "Argument or function not available")


"""
    guess_datatype(fmtcode::AbstractString)
guess `DataType` from Excel format code
"""
function guess_datatype(fmtcode::AbstractString)
    T = String
    for reg in ((ExcelDateTime,    r"[ymdhs]"),
                (ScientificFormat, r"E\+|E\-"),
                (Float64,          r"0\.|#\.|0_|#_|\%|\?\/"),
                (Float64,          r"0|#"))
        if ismatch(reg[2], fmtcode)
            T = reg[1]
            break
        end
    end

    # T == String && warn("failed to find 'DataType' for '$fmtcode'")
    return T
end
