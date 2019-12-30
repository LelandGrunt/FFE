@echo off

:: https://benchmarkdotnet.org/articles/guides/console-args.html

.\FFE-AddIn_Benchmark.exe --f FFE.WebParserPerformance*
.\FFE-AddIn_Benchmark.exe --f FFE.WebParserPerformance* --allCategories %1
.\FFE-AddIn_Benchmark.exe --f FFE.WebParserPerformanceWeb*
.\FFE-AddIn_Benchmark.exe --f FFE.WebParserMethodExpressionPerformance*
.\FFE-AddIn_Benchmark.exe --f FFE.WebParserMethodExpressionPerformanceWeb*
.\FFE-AddIn_Benchmark.exe --f FFE.WebParserMethodExpressionPerformanceWeb* --launchCount 1 --iterationCount 100

pause