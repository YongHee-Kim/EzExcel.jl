EzXML.parsexml(x::ReadableFile) = parsexml(readstring(x))
EzXML.parsexml(x::Reader) = parsexml(readstring(x))

"""
    ScientificFormat

Scientific notation numbers in Excel.
If the format code is `0.00E+00` and the cell value is 12,200,000, then 1.22E+07 is
stored in the xml.
"""
struct ScientificFormat end
function Base.parse(::Type{ScientificFormat}, x::AbstractString)::Float64
    if contains(x, "E")
        x = lowercase(x)
    end
    parse(Float64, x)
end

"""
    GeneralFormat

The primary goal when a cell is using `GeneralFormat` is to render the cell content without user-specified
guidance to the best ability of the application.
It can be `String`, `Boolean`, `Error`, `Numbers`
`DateTime` do not fall into "GeneralFormat`
"""
struct GeneralFormat end
function Base.parse(::Type{GeneralFormat}, x)
    function isfloat(x)::Bool
        b = false
        if length(x) > 1 
            if contains(x, ".")
                b = length(matchall(r"[.]", x)) < 2 ? true : false
            end
            if startswith(x, "-")
                b = length(matchall(r"[-]", x)) < 2 ? true : false
            end
        end
        b
    end
    function isint(x::AbstractString)::Bool
        b = false
        if !isempty(x)
            cha = matchall(r"[^0-9]", x)
            b = isempty(cha)                         ? true : 
                (length(cha) == 1 && cha[1] == "-" ) ? true : 
                                                     false
        end
        b
    end
    if isfloat(x)
        parse(Float64, x)
    else
        in(x, keys(EzExcel.ECMA_ERROR_VALUES))    ? missing :
        in(x, ("True", "true", "False", "false")) ? parse(Bool, lowercase(x)) :
        isint(x)                                  ? parse(Int, x) :
                                                  String(x)
    end
end    


"""
    ExcelDateTime

Parse an Excel number (presumed to represent a date, a datetime or a time) to a DateTime

xldate: The Excel number
datemode: true:1904-based, false: 1900-based.

    `1904-01-01` is not regarded as a valid date in the `datemode == true`
     its "serial number" is zero.
 """
struct ExcelDateTime end
function Base.parse(::Type{ExcelDateTime}, xldate::AbstractString; args...)
    parse(ExcelDateTime, parse(ScientificFormat, xldate); args...)
end
function Base.parse(::Type{ExcelDateTime}, xldate::Real; datemode::Bool=false)::DateTime
    # Set the epoch based on the 1900/1904 datemode.
    epoch = if datemode
        DateTime(1904, 1, 1)
    else # Workaround Excel 1900 leap year bug by adjusting the epoch.
        xldate < 60 ? DateTime(1899, 12, 31) : DateTime(1899, 12, 30)
    end
    xldays, fraction = divrem(xldate, 1)
    xldays == 0 && (xldays = 1)

    sec = round(Int, fraction * 86400000.0)
    sec, milli_sec = divrem(sec, 1000)

    return epoch + Dates.Day(xldays) + Dates.Second(sec) + Dates.Millisecond(milli_sec)
end

struct SharedString end
