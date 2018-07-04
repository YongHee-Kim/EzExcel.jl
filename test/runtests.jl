using EzExcel

@static if VERSION < v"0.7.0-DEV.2005"
    using Base.Test
    using Missings
else
    using Test
end

PATH = Pkg.dir("EzExcel") |> x -> joinpath(x, "test\\datasets")

# testset.xlsx
filename = joinpath(PATH, "testset.xlsx")
wb = WorkBook(filename)

@test length(wb) == 4
@test sheetname(wb) == ("combination", "formats", "empty", "negative")

wss = sheets(wb)

@testset "testset:[combination]" begin
    ws = wss[1]
    sheet_1 = [["Some Numbers", 1, 1.5, 2, 2.5], 
               ["Some Strings", "A", "BB", "CCC", "DDDD"],
               ["Some Bools", true, false, false, true],
               ["Mixed column", 2, "EEEEE", false, 1.5],
               ["Mixed with NA", 9, "III", missing, true],
               ["Float64 with NA", 3, missing, 3.5, 4],
               ["String with NA", "FF", missing, "GGG", "HHHH"],
               ["Bool with NA", missing, true, missing, false],
               ["Some dates", DateTime(2015, 3, 3), DateTime(2015, 2, 4, 10, 14), DateTime(1988, 4, 9, 0, 00), DateTime(1900, 1, 1, 15, 02, 00)],
               ["Some errors", missing, missing, missing, missing],
               ["Errors with NA", missing, missing, missing, missing]]
    sheet_1 = hcat(sheet_1...)

    @test sheetname(ws) == "combination"

     for r in 1:size(ws, 1), c in 1:size(ws, 2)
        if ismissing(ws[r, c])
            @test ismissing(sheet_1[r, c])
        else
            @test ws[r, c] == sheet_1[r, c]
        end
    end

end

filename = joinpath(PATH, "big_testset.xlsx")

