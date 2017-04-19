.\packages\OpenCover.4.6.519\tools\OpenCover.Console.exe `
    -register:user -target:"packages\xunit.runner.console.2.2.0\tools\xunit.console.exe" `
    -targetargs:".\UnitTests\bin\Debug\UnitTests.dll" `
    -output:"coverage.xml" -searchdirs:".\UnitTests\bin\Debug" 

.\packages\ReportGenerator.2.5.6\tools\ReportGenerator.exe `
    -reports:"coverage.xml" `
    -targetdir:".\coverage\"