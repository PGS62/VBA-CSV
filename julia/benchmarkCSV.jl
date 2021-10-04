using CSV#= as of 29 Aug 2021 one needs the main version of CSV, latest released version
0.8.5 fails to load some of the files. See https://github.com/JuliaData/CSV.jl/issues/879
Use:
]add CSV#main
=#
using DataFrames

function benchmark()
    #= folder points to where the VBA method RunSpeedTests writes performance test files.
    See workbooks/VBA-CSV.xlsm, module modCSVPerformance =# 

    folder = joinpath(ENV["TEMP"],"VBA-CSV/Performance")
    outputfile = normpath(joinpath(@__DIR__,"..","julia/juliaparsetimes.csv"))
 outputfile = outputfile * "2"
    benchmark_csvs_in_folder(folder,outputfile)
end

"""
   benchmark_csvs_in_folder(folder::String, outputfile::String)
Benchmark all .csv files in `folder`, writing results to `outputfile`.
"""
function benchmark_csvs_in_folder(folder::String, outputfile::String)
    files = readdir(folder, join=true)
    files = filter(x -> x[end - 3:end] == ".csv", files)# only .csv files

    n = length(files)
    times = fill(0.0, n)
    numcalls = fill(0, n)
    statuses = fill("OK", n)

    foo = benchmarkonefile(files[1], 1)# for compilation "warmup"
    for (f, i) in zip(files, 1:n)
        println(i, f)
        try
            times[i], numcalls[i] = benchmarkonefile(f, 5)
        catch e
            statuses[i] = "$e"
        end
    end
    times

    result = DataFrame(filename=replace.(files, "/" => "\\"), time=times, 
                        status=statuses, numcalls=numcalls)
    CSV.write(outputfile, result)
end

"""
    benchmarkonefile(filename::String, timeout::Int)
Average time (over sufficient trials to take `timeout` seconds) to load file `filename` to
a DataFrame, using CSV.File.
"""
function benchmarkonefile(filename::String, timeout::Int)
    i = 0 ; time2 = time() # needed to give variables scope outside the loop.
    time1 = time()
    while true
        i = i + 1
        res = CSV.File(filename, header=false, type=String) |> DataFrame
        time2 = time()
        time2 - time1 < timeout || break
    end
    (time2 - time1) / i, i
end

function shownotcompliantfile(filenum::String)
filename = "C:/Projects/VBA-CSV/testfiles/Not_RFC_4180_Compliant_" * filenum * ".csv"
res = CSV.File(filename, header=false, delim = ",",type=String) |>DataFrame
Matrix(res)
end