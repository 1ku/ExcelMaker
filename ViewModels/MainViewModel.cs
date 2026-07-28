﻿using Microsoft.Win32;

namespace ExcelMaker.ViewModels;

public partial class MainViewModel : ObservableObject
{
    private readonly DatabaseService _db;
    private readonly ConfigService _config;

    [ObservableProperty] private string _username = string.Empty;   // 工号
    [ObservableProperty] private string _currentUser = string.Empty; // 全名
    [ObservableProperty] private string _statusMessage = "就绪";
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private DateTime _selectedYearMonth = DateTime.Today; // 年月，默认当天

    public ObservableCollection<string> LogEntries { get; } = new();

    public MainViewModel(DatabaseService db, ConfigService config)
    {
        _db = db;
        _config = config;
    }

    private void AddLog(string msg)
    {
        var line = $"{DateTime.Now:HH:mm:ss}  {msg}";
        Application.Current.Dispatcher.Invoke(() => LogEntries.Add(line));
        Log.Information(msg);
    }

    #region 功能2：导入更新库位名称

    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task ImportStockNameAsync()
    {
        var dlg = new OpenFileDialog
        {
            Title = "选择库位名称 Excel",
            Filter = "Excel 文件|*.xlsx;*.xls|所有文件|*.*"
        };
        if (dlg.ShowDialog() != true) return;

        IsBusy = true;
        StatusMessage = "正在导入库位名称…";
        try
        {
            AddLog($"读取 Excel：{dlg.FileName}");
            var rows = ExcelHelper.ReadStockNameSheet(dlg.FileName);
            AddLog($"识别到 {rows.Count} 条库位数据");

            if (rows.Count == 0)
            {
                StatusMessage = "未读取到有效数据行";
                return;
            }

            var n = await _db.MergeStockNameAsync(rows, Username);
            AddLog($"MERGE 完成，处理 {n} 行（AUTHOR={Username}）");
            StatusMessage = $"库位名称更新成功：{n} 行";
        }
        catch (Exception ex)
        {
            Log.Error(ex, "导入库位名称失败");
            AddLog($"失败：{ex.Message}");
            StatusMessage = $"失败：{ex.Message}";
        }
        finally { IsBusy = false; }
    }

    #endregion

    #region 功能3：处理库存明细

    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task ProcessInventoryAsync()
    {
        var openDlg = new OpenFileDialog
        {
            Title = "选择库存明细 Excel（A 表）",
            Filter = "Excel 文件|*.xlsx;*.xls|所有文件|*.*"
        };
        if (openDlg.ShowDialog() != true) return;

        var saveDlg = new SaveFileDialog
        {
            Title = "保存处理后的库存明细（B 表）",
            Filter = "Excel 文件|*.xlsx",
            FileName = $"库存明细_{SelectedYearMonth:yyyyMM}.xlsx",
            InitialDirectory = _config.OutputDir
        };
        if (saveDlg.ShowDialog() != true) return;

        var ym = SelectedYearMonth.ToString("yyyyMM");
        IsBusy = true;
        StatusMessage = "正在处理库存明细…";
        try
        {
            AddLog($"读取 A 表：{openDlg.FileName}");
            var aRows = ExcelHelper.ReadInventorySheet(openDlg.FileName);
            AddLog($"识别到 {aRows.Count} 条明细");

            if (!File.Exists(_config.TemplatePath))
                throw new FileNotFoundException($"未找到库存明细模板：{_config.TemplatePath}", _config.TemplatePath);

            AddLog("预加载库位名称字典（FACTORY_STOCK_NAME）…");
            var stockDict = await _db.GetStockNameLookupAsync();

            AddLog($"预加载复盘人字典（D_S_INVENTORY_DETAIL，年月={ym}）…");
            var authorDict = await _db.GetInventoryAuthorLookupAsync(ym);

            AddLog($"基于模板生成：{_config.TemplatePath}");
            ExcelHelper.BuildInventoryWorkbook(_config.TemplatePath, saveDlg.FileName, aRows, stockDict, authorDict);

            AddLog($"已导出：{saveDlg.FileName}");
            StatusMessage = $"库存明细处理完成：{aRows.Count} 行 → {Path.GetFileName(saveDlg.FileName)}";
        }
        catch (Exception ex)
        {
            Log.Error(ex, "处理库存明细失败");
            AddLog($"失败：{ex.Message}");
            StatusMessage = $"失败：{ex.Message}";
        }
        finally { IsBusy = false; }
    }

    #endregion

    private bool CanRun() => !IsBusy;

    [RelayCommand]
    private void ExitApp() => Application.Current.Shutdown();
}
