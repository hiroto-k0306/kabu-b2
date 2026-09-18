// 東証プライム全銘柄の長期日足から、README 15. の信号ごとに「上位10を等分」した夜間リターンを日ごとに計算する。
// PowerShell の Add-Type でコンパイルして使う（Windows 標準の .NET だけで動くよう C# 5 の文法に限定）。
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

public class CrossSectionStudy
{
    public const double MaxValidPrice = 5000000.0;
    public const double MinOvernightRatio = 0.4;
    public const double MaxOvernightRatio = 2.5;
    public const int MaxDateGapDays = 30;

    public static readonly string[] SignalNames = new string[] { "A1", "A2", "A3", "A4", "B1", "B2", "B3", "V1", "B3L", "B3VR", "B3LS" };
    // B3VR（README 24. A1）: B3 ÷ 値動き の上位。B3LS（24. A2）: B3L の順位で同じ業種は2銘柄まで
    // SectorOf: 銘柄コード → 業種番号（B3LS で使う。null なら業種の制限なし）
    public static Dictionary<string, int> SectorOf = null;
    // V1 は事前登録外の点検用: 直近250日の日次（終値→終値）対数リターンの標準偏差が大きい10銘柄
    // B3L（README 22. B2）: その日の対象銘柄のうち V1 の値動きが中央値以下の銘柄に絞った B3

    public class DayRow
    {
        public string Date;          // 判断日 t
        public string BuyDate;       // t+1（引けで買う）
        public string SellDate;      // t+2（寄りで売る）
        public int Eligible;
        public double MarketEw;      // 対象銘柄の等分の夜間リターン
        public double N225;          // 日経平均の夜間リターン（配当なし）
        public double[] Signal;      // SignalNames の順。上位10が揃わない日は NaN
        public string[] Picks;
        // 各信号の上位topKの、順位ごとの夜間リターン（README 25.）
        public double[][] PickRets;
        // RunWithGroups のときだけ: グループ番号(1,2)ごとの 全銘柄等分 / 平均売買代金上位 topK の等分 のリターン
        public double[] GroupEw;
        public double[] GroupTop;
        public int[] GroupCount;
        public string[] GroupPicks;
    }

    static double ParseD(string s)
    {
        s = s.Trim('"');
        if (s.Length == 0) return double.NaN;
        // 壊れたデータ（"-∞" など）は NaN にして不正な行として扱う
        double v;
        if (!double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out v)) return double.NaN;
        return v;
    }

    public static List<DayRow> Run(string tickerDir, string n225Csv, double minTurnover, int minEligible, int topK, string logPath, double minRawPrice)
    {
        return RunWithGroups(tickerDir, n225Csv, minTurnover, minEligible, topK, logPath, minRawPrice, null);
    }

    // groupOf: 銘柄コード → グループ番号（1 または 2）。null ならグループ集計をしない
    public static List<DayRow> RunWithGroups(string tickerDir, string n225Csv, double minTurnover, int minEligible, int topK, string logPath, double minRawPrice, Dictionary<string, int> groupOf)
    {
        // --- 暦（日経平均の営業日） ---
        var calDates = new List<string>();
        var n225Open = new List<double>();
        var n225Close = new List<double>();
        bool header = true;
        foreach (var line in File.ReadLines(n225Csv))
        {
            if (header) { header = false; continue; }
            var f = line.Split(',');
            calDates.Add(f[0].Trim('"'));
            n225Open.Add(ParseD(f[2]));
            n225Close.Add(ParseD(f[5]));
        }
        int T = calDates.Count;
        var calIdx = new Dictionary<string, int>();
        for (int i = 0; i < T; i++) calIdx[calDates[i]] = i;
        var calDt = new DateTime[T];
        for (int i = 0; i < T; i++) calDt[i] = DateTime.ParseExact(calDates[i], "yyyy-MM-dd", CultureInfo.InvariantCulture);

        // --- 銘柄 ---
        var files = Directory.GetFiles(tickerDir, "*.csv");
        Array.Sort(files, StringComparer.Ordinal);
        int S = files.Length;
        var codes = new string[S];
        var aO = new float[S][];
        var aC = new float[S][];
        var vol = new float[S][];
        var rawC = new float[S][];
        var ok = new bool[S][];
        long markedTotal = 0;
        using (var log = new StreamWriter(logPath))
        {
            for (int s = 0; s < S; s++)
            {
                codes[s] = Path.GetFileNameWithoutExtension(files[s]);
                aO[s] = new float[T]; aC[s] = new float[T]; vol[s] = new float[T]; rawC[s] = new float[T]; ok[s] = new bool[T];
                var rowDates = new List<DateTime>();
                var rowCal = new List<int>();
                var ro = new List<double>(); var rc = new List<double>(); var rrc = new List<double>(); var rv = new List<double>();
                header = true;
                foreach (var line in File.ReadLines(files[s]))
                {
                    if (header) { header = false; continue; }
                    var f = line.Split(',');
                    if (f.Length < 11) continue;
                    string d = f[0].Trim('"');
                    rowDates.Add(DateTime.ParseExact(d, "yyyy-MM-dd", CultureInfo.InvariantCulture));
                    int ci;
                    rowCal.Add(calIdx.TryGetValue(d, out ci) ? ci : -1);
                    ro.Add(ParseD(f[2])); rc.Add(ParseD(f[5])); rv.Add(ParseD(f[6])); rrc.Add(ParseD(f[10]));
                }
                // Set-InvalidPriceRows と同じ基準（出来高0の行も不正として扱う）
                int lastValid = -1, marked = 0;
                for (int i = 0; i < ro.Count; i++)
                {
                    double o = ro[i], c = rc[i], raw = rrc[i], v = rv[i];
                    bool valid = o > 0 && c > 0 && o <= MaxValidPrice && c <= MaxValidPrice
                        && (double.IsNaN(raw) || (raw >= 1 && raw <= MaxValidPrice)) && v > 0;
                    bool isBreak = false;
                    if (valid && lastValid >= 0)
                    {
                        if ((rowDates[i] - rowDates[lastValid]).TotalDays > MaxDateGapDays) isBreak = true;
                        else if (lastValid == i - 1)
                        {
                            double ratio = o / rc[lastValid];
                            if (ratio > MaxOvernightRatio || ratio < MinOvernightRatio) isBreak = true;
                        }
                    }
                    if (valid && !isBreak)
                    {
                        lastValid = i;
                        int ci = rowCal[i];
                        if (ci >= 0) { aO[s][ci] = (float)o; aC[s][ci] = (float)c; vol[s][ci] = (float)v; rawC[s][ci] = (float)raw; ok[s][ci] = true; }
                        continue;
                    }
                    if (isBreak) lastValid = -1;
                    marked++;
                }
                markedTotal += marked;
                if (marked > 0) log.WriteLine(codes[s] + "\t" + marked + " rows invalid");
            }
            log.WriteLine("total invalid rows: " + markedTotal);
        }

        // --- 日ごとの計算（窓の合計は日を進めながら足し引きする） ---
        // 窓 [t-a, t-b] を持つ合計。value(s, i) が NaN の日は数えない
        var to20prev = new Roll(S, 20, 1);   // t-20..t-1
        var to20incl = new Roll(S, 19, 0);   // t-19..t
        var to5 = new Roll(S, 4, 0);         // t-4..t
        var to20b = new Roll(S, 24, 5);      // t-24..t-5
        var on250 = new Roll(S, 249, 0);
        var on60 = new Roll(S, 59, 0);
        var id250 = new Roll(S, 249, 0);
        var cc250 = new Roll(S, 249, 0);
        var cc250sq = new Roll(S, 249, 0);

        Func<int, int, double> TO = (s, i) => ok[s][i] ? (double)aC[s][i] * vol[s][i] : double.NaN;
        Func<int, int, double> ON = (s, i) => (i > 0 && ok[s][i] && ok[s][i - 1]) ? Math.Log((double)aO[s][i] / aC[s][i - 1]) : double.NaN;
        Func<int, int, double> CC = (s, i) => (i > 0 && ok[s][i] && ok[s][i - 1]) ? Math.Log((double)aC[s][i] / aC[s][i - 1]) : double.NaN;
        Func<int, int, double> CC2 = (s, i) => { double v = CC(s, i); return v * v; };
        Func<int, int, double> ID = (s, i) => ok[s][i] ? Math.Log((double)aC[s][i] / aO[s][i]) : double.NaN;

        var rows = new List<DayRow>();
        int K = SignalNames.Length;
        var scores = new List<KeyValuePair<double, int>>[K];
        for (int k = 0; k < K; k++) scores[k] = new List<KeyValuePair<double, int>>(S);
        var rets = new double[S];
        var volOf = new double[S];
        var b3Of = new double[S];
        var groupIdx = new int[S];
        if (groupOf != null) { for (int s = 0; s < S; s++) { int g; groupIdx[s] = groupOf.TryGetValue(codes[s], out g) ? g : 0; } }
        var gLists = new List<KeyValuePair<double, int>>[3];
        for (int g = 0; g < 3; g++) gLists[g] = new List<KeyValuePair<double, int>>();
        var gSum = new double[3];
        var gCnt = new int[3];

        for (int t = 0; t < T; t++)
        {
            to20prev.Step(t, TO); to20incl.Step(t, TO); to5.Step(t, TO); to20b.Step(t, TO);
            on250.Step(t, ON); on60.Step(t, ON); id250.Step(t, ID); cc250.Step(t, CC); cc250sq.Step(t, CC2);
            if (t + 2 >= T) continue;

            for (int k = 0; k < K; k++) scores[k].Clear();
            int eligible = 0; double sumRet = 0;
            for (int g = 0; g < 3; g++) { gLists[g].Clear(); gSum[g] = 0; gCnt[g] = 0; }
            for (int s = 0; s < S; s++)
            {
                if (!ok[s][t] || !ok[s][t + 1] || !ok[s][t + 2]) continue;
                if (to20incl.Cnt[s] < 15 || to20incl.Sum[s] / to20incl.Cnt[s] < minTurnover) continue;
                if (minRawPrice > 0 && !(rawC[s][t] >= minRawPrice)) continue;
                double ret = (double)aO[s][t + 2] / aC[s][t + 1] - 1;
                rets[s] = ret;
                eligible++; sumRet += ret;
                int gi = groupIdx[s];
                if (gi > 0) { gSum[gi] += ret; gCnt[gi]++; gLists[gi].Add(new KeyValuePair<double, int>(to20incl.Sum[s] / to20incl.Cnt[s], s)); }

                double toT = TO(s, t);
                double a1 = double.NaN;
                if (to20prev.Cnt[s] >= 15 && to20prev.Sum[s] > 0) a1 = toT / (to20prev.Sum[s] / to20prev.Cnt[s]);
                if (!double.IsNaN(a1))
                {
                    scores[0].Add(new KeyValuePair<double, int>(a1, s));
                    if (t > 0 && ok[s][t - 1])
                    {
                        if (aC[s][t] > aC[s][t - 1]) scores[2].Add(new KeyValuePair<double, int>(a1, s));
                        else if (aC[s][t] < aC[s][t - 1]) scores[3].Add(new KeyValuePair<double, int>(a1, s));
                    }
                }
                if (to5.Cnt[s] >= 4 && to20b.Cnt[s] >= 15 && to20b.Sum[s] > 0)
                    scores[1].Add(new KeyValuePair<double, int>((to5.Sum[s] / to5.Cnt[s]) / (to20b.Sum[s] / to20b.Cnt[s]), s));
                if (on250.Cnt[s] >= 200)
                {
                    scores[4].Add(new KeyValuePair<double, int>(on250.Sum[s], s));
                    b3Of[s] = double.NaN;
                    if (id250.Cnt[s] >= 200) { b3Of[s] = on250.Sum[s] - id250.Sum[s]; scores[6].Add(new KeyValuePair<double, int>(b3Of[s], s)); }
                }
                if (cc250.Cnt[s] >= 200)
                {
                    double m = cc250.Sum[s] / cc250.Cnt[s];
                    volOf[s] = Math.Sqrt(Math.Max(0, cc250sq.Sum[s] / cc250.Cnt[s] - m * m));
                    scores[7].Add(new KeyValuePair<double, int>(volOf[s], s));
                }
                if (on60.Cnt[s] >= 48) scores[5].Add(new KeyValuePair<double, int>(on60.Sum[s], s));
            }
            if (eligible < minEligible) continue;
            if (scores[7].Count > 0)
            {
                var vols = new double[scores[7].Count];
                for (int j = 0; j < vols.Length; j++) vols[j] = scores[7][j].Key;
                Array.Sort(vols);
                double median = (vols.Length % 2 == 1) ? vols[vols.Length / 2] : (vols[vols.Length / 2 - 1] + vols[vols.Length / 2]) / 2;
                foreach (var kv in scores[7])
                {
                    int s = kv.Value;
                    if (on250.Cnt[s] >= 200 && id250.Cnt[s] >= 200)
                    {
                        if (kv.Key <= median) { scores[8].Add(new KeyValuePair<double, int>(b3Of[s], s)); scores[10].Add(new KeyValuePair<double, int>(b3Of[s], s)); }
                        if (kv.Key > 0) scores[9].Add(new KeyValuePair<double, int>(b3Of[s] / kv.Key, s));
                    }
                }
            }

            var row = new DayRow();
            row.Date = calDates[t]; row.BuyDate = calDates[t + 1]; row.SellDate = calDates[t + 2];
            row.Eligible = eligible;
            row.MarketEw = sumRet / eligible;
            row.N225 = n225Open[t + 2] / n225Close[t + 1] - 1;
            row.Signal = new double[K];
            row.Picks = new string[K];
            row.PickRets = new double[K][];
            for (int k = 0; k < K; k++)
            {
                var list = scores[k];
                if (list.Count < topK) { row.Signal[k] = double.NaN; row.Picks[k] = ""; row.PickRets[k] = null; continue; }
                list.Sort(delegate (KeyValuePair<double, int> x, KeyValuePair<double, int> y)
                {
                    int c = y.Key.CompareTo(x.Key);
                    return c != 0 ? c : x.Value.CompareTo(y.Value);
                });
                double sr = 0; var picks = new string[topK];
                var chosen = new int[topK];
                if (k == 10 && SectorOf != null)
                {
                    // 同じ業種は2銘柄まで
                    var perSector = new Dictionary<int, int>();
                    int got = 0;
                    for (int j = 0; j < list.Count && got < topK; j++)
                    {
                        int sec; if (!SectorOf.TryGetValue(codes[list[j].Value], out sec)) sec = -1 - j;
                        int used; perSector.TryGetValue(sec, out used);
                        if (used >= 2) continue;
                        perSector[sec] = used + 1;
                        sr += rets[list[j].Value]; picks[got] = codes[list[j].Value]; chosen[got] = list[j].Value; got++;
                    }
                    if (got < topK) { row.Signal[k] = double.NaN; row.Picks[k] = ""; continue; }
                }
                else
                {
                    for (int j = 0; j < topK; j++) { sr += rets[list[j].Value]; picks[j] = codes[list[j].Value]; chosen[j] = list[j].Value; }
                }
                row.Signal[k] = sr / topK;
                row.Picks[k] = string.Join(" ", picks);
                var pr = new double[topK];
                for (int j = 0; j < topK; j++) pr[j] = rets[chosen[j]];
                row.PickRets[k] = pr;
            }
            if (groupOf != null)
            {
                row.GroupEw = new double[3]; row.GroupTop = new double[3]; row.GroupCount = new int[3]; row.GroupPicks = new string[3];
                for (int g = 1; g < 3; g++)
                {
                    row.GroupCount[g] = gCnt[g];
                    row.GroupEw[g] = gCnt[g] > 0 ? gSum[g] / gCnt[g] : double.NaN;
                    var list = gLists[g];
                    if (list.Count < topK) { row.GroupTop[g] = double.NaN; row.GroupPicks[g] = ""; continue; }
                    list.Sort(delegate (KeyValuePair<double, int> x, KeyValuePair<double, int> y)
                    {
                        int c = y.Key.CompareTo(x.Key);
                        return c != 0 ? c : x.Value.CompareTo(y.Value);
                    });
                    double sr = 0; var picks = new string[topK];
                    for (int j = 0; j < topK; j++) { sr += rets[list[j].Value]; picks[j] = codes[list[j].Value]; }
                    row.GroupTop[g] = sr / topK;
                    row.GroupPicks[g] = string.Join(" ", picks);
                }
            }
            rows.Add(row);
        }
        return rows;
    }

    public class Roll
    {
        public double[] Sum;
        public int[] Cnt;
        int a, b;
        public Roll(int n, int a, int b) { Sum = new double[n]; Cnt = new int[n]; this.a = a; this.b = b; }
        public void Step(int t, Func<int, int, double> value)
        {
            int add = t - b, rem = t - a - 1;
            for (int s = 0; s < Sum.Length; s++)
            {
                if (add >= 0) { double v = value(s, add); if (!double.IsNaN(v) && !double.IsInfinity(v)) { Sum[s] += v; Cnt[s]++; } }
                if (rem >= 0) { double v = value(s, rem); if (!double.IsNaN(v) && !double.IsInfinity(v)) { Sum[s] -= v; Cnt[s]--; } }
            }
        }
    }
}
