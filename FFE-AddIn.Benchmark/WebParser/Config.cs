using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Engines;
using BenchmarkDotNet.Environments;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Order;
using BenchmarkDotNet.Reports;
using BenchmarkDotNet.Running;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;

namespace FFE
{
    public class Config : ManualConfig
    {
        public Config()
        {
            Add(new ProviderColumn());
            Orderer = new WebParserOrderer();
        }
    }

    public class WebConfig : Config
    {
        public WebConfig() : base()
        {
            Job jobX86 = new Job("WebTestsJobX86")
            {
                Environment =
                {
                    Platform = Platform.X86,
                    Runtime = ClrRuntime.Net462
                },
                Run =
                {
                    RunStrategy = RunStrategy.ColdStart,
                    LaunchCount = 1,
                    IterationCount = 100
                }
            };

            Job jobX64 = new Job("WebTestsJobX64", jobX86).With(Platform.X64);

            Add(jobX86);
            Add(jobX64);
        }
    }

    public class ProviderColumn : IColumn
    {
        public ProviderColumn() { }

        public string Id => nameof(ProviderColumn);
        public string ColumnName => "Provider";
        public bool AlwaysShow => true;
        public ColumnCategory Category => ColumnCategory.Params;
        public int PriorityInCategory => 0;
        public bool IsNumeric => false;
        public UnitType UnitType => UnitType.Dimensionless;
        public string Legend => "Name of Provider";
        public string GetValue(Summary summary, BenchmarkCase benchmarkCase)
        {
            string provider = ((Provider)benchmarkCase.Parameters.GetArgument("Provider").Value).Name;
            return provider;
        }
        public string GetValue(Summary summary, BenchmarkCase benchmarkCase, SummaryStyle style) => GetValue(summary, benchmarkCase);
        public bool IsAvailable(Summary summary) => true;
        public bool IsDefault(Summary summary, BenchmarkCase benchmarkCase) => false;
    }

    public class WebParserOrderer : IOrderer
    {
        public bool SeparateLogicalGroups => true;

        public IEnumerable<BenchmarkCase> GetExecutionOrder(ImmutableArray<BenchmarkCase> benchmarksCase) =>
            from benchmark in benchmarksCase
            orderby ((Provider)benchmark.Parameters.GetArgument("Provider").Value).Name ascending
            select benchmark;

        public string GetHighlightGroupKey(BenchmarkCase benchmarkCase) => null;

        public string GetLogicalGroupKey(ImmutableArray<BenchmarkCase> allBenchmarksCases, BenchmarkCase benchmarkCase) =>
            ((Provider)benchmarkCase.Parameters.GetArgument("Provider").Value).Name;

        public IEnumerable<IGrouping<string, BenchmarkCase>> GetLogicalGroupOrder(IEnumerable<IGrouping<string, BenchmarkCase>> logicalGroups) =>
            logicalGroups.OrderBy(lg => lg.Key);

        public IEnumerable<BenchmarkCase> GetSummaryOrder(ImmutableArray<BenchmarkCase> benchmarksCases, Summary summary) =>
            from benchmarkCase in benchmarksCases
            orderby ((Provider)benchmarkCase.Parameters.GetArgument("Provider").Value).Name ascending,
                    summary[benchmarkCase].ResultStatistics?.Mean ascending
            select benchmarkCase;
    }
}